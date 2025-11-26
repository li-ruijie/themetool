#!/usr/bin/env luajit
--[[
    ThemeTool - Windows Theme Manager for LuaJIT

    This tool provides a command-line interface for managing Windows themes.
    It uses LuaJIT's FFI (Foreign Function Interface) to interact with:
    - UxTheme.dll: Windows theme API
    - Ole32.dll: COM (Component Object Model) infrastructure
    - OleAut32.dll: OLE Automation for BSTR string management

    The tool interfaces with undocumented Windows COM interfaces:
    - IThemeManager: Manages Windows themes
    - ITheme: Represents a single theme with its properties

    Commands supported:
    - getthemestatus: Check if themes are active
    - getcurrentthemename: Get the display name of current theme
    - getcurrentvisualstylename: Get the visual style filename
    - changetheme: Apply a new theme from a .theme file
]]

local ffi = require("ffi")

--[[
    ============================================================================
    SECTION 1: FFI C DECLARATIONS
    ============================================================================
    This section declares all the C structures and functions we'll use from
    Windows DLLs. LuaJIT's FFI allows us to call these directly without
    writing any C wrapper code.
]]

ffi.cdef[[
    // ========================================================================
    // UxTheme.dll - Windows Theme API
    // ========================================================================

    // Check if Windows themes/visual styles are currently enabled
    // Returns: true if themes are active, false otherwise
    bool IsThemeActive(void);

    // ========================================================================
    // Ole32.dll - Component Object Model (COM) Support
    // ========================================================================

    // GUID (Globally Unique Identifier) structure - 128-bit identifier
    // Used to identify COM interfaces (IID) and classes (CLSID)
    typedef struct _GUID {
        uint32_t Data1;          // First 4 bytes
        uint16_t Data2;          // Next 2 bytes
        uint16_t Data3;          // Next 2 bytes
        uint8_t  Data4[8];       // Final 8 bytes
    } GUID;

    // Type aliases for clarity
    typedef GUID IID;            // Interface ID (identifies COM interfaces)
    typedef GUID CLSID;          // Class ID (identifies COM classes/components)
    typedef void* LPVOID;        // Generic pointer type
    typedef int32_t HRESULT;     // COM function return code (0 = success, negative = error)
    typedef uint16_t OLECHAR;    // Wide character (UTF-16)
    typedef OLECHAR* BSTR;       // BSTR = Basic String (COM string type with length prefix)

    // Initialize the COM library for the current thread
    // pvReserved: Must be NULL
    // Returns: S_OK (0) on success, or error code
    HRESULT CoInitialize(void* pvReserved);

    // Uninitialize the COM library (cleanup)
    void CoUninitialize(void);

    // Create an instance of a COM object
    // rclsid: Class ID of the object to create
    // pUnkOuter: NULL for non-aggregated objects
    // dwClsContext: Execution context (CLSCTX_INPROC_SERVER = same process)
    // riid: Interface ID we want to use
    // ppv: Output pointer to receive the interface
    HRESULT CoCreateInstance(
        const CLSID* rclsid,
        void* pUnkOuter,
        uint32_t dwClsContext,
        const IID* riid,
        LPVOID* ppv
    );

    // ========================================================================
    // OleAut32.dll - BSTR String Management
    // ========================================================================
    // BSTR is a COM string type with a length prefix stored before the string.
    // Must be allocated/freed with these functions (not malloc/free!)

    // Allocate a new BSTR from a wide string
    BSTR SysAllocString(const OLECHAR* psz);

    // Free a BSTR (deallocate memory)
    void SysFreeString(BSTR bstrString);

    // Get the length of a BSTR (in characters, not including null terminator)
    uint32_t SysStringLen(BSTR bstrString);

    // ========================================================================
    // IUnknown - Base interface for all COM objects
    // ========================================================================
    // Every COM interface inherits from IUnknown, which provides:
    // - QueryInterface: Get other interfaces from the same object
    // - AddRef: Increment reference count
    // - Release: Decrement reference count (object deletes itself when count reaches 0)

    typedef struct IUnknownVtbl {
        HRESULT (*QueryInterface)(void* This, const IID* riid, void** ppvObject);
        uint32_t (*AddRef)(void* This);
        uint32_t (*Release)(void* This);
    } IUnknownVtbl;

    // ========================================================================
    // Kernel32.dll - String Conversion Functions
    // ========================================================================
    // Windows uses UTF-16 (wide chars) internally, but Lua uses UTF-8.
    // These functions convert between the two encodings.

    // Convert multibyte string (UTF-8) to wide char string (UTF-16)
    // CodePage: Character encoding (CP_UTF8 = 65001)
    // dwFlags: Conversion flags (0 = default behavior)
    // lpMultiByteStr: Input UTF-8 string
    // cbMultiByte: Length of input in bytes (-1 = null-terminated)
    // lpWideCharStr: Output buffer for UTF-16 (NULL to query required size)
    // cchWideChar: Size of output buffer in characters
    // Returns: Number of characters written (or required if lpWideCharStr is NULL)
    int MultiByteToWideChar(
        uint32_t CodePage,
        uint32_t dwFlags,
        const char* lpMultiByteStr,
        int cbMultiByte,
        OLECHAR* lpWideCharStr,
        int cchWideChar
    );

    // Convert wide char string (UTF-16) to multibyte string (UTF-8)
    // Parameters similar to MultiByteToWideChar but in reverse
    int WideCharToMultiByte(
        uint32_t CodePage,
        uint32_t dwFlags,
        const OLECHAR* lpWideCharStr,
        int cchWideChar,
        char* lpMultiByteStr,
        int cbMultiByte,
        const char* lpDefaultChar,
        int* lpUsedDefaultChar
    );
]]

--[[
    ============================================================================
    SECTION 2: LOAD WINDOWS DLLS
    ============================================================================
    Load the Windows DLLs we'll be calling through FFI.
]]

local uxtheme = ffi.load("uxtheme.dll")    -- Windows theme/visual style API
local ole32 = ffi.load("ole32.dll")        -- COM infrastructure
local oleaut32 = ffi.load("oleaut32.dll")  -- OLE Automation (BSTR functions)

--[[
    ============================================================================
    SECTION 3: CONSTANTS
    ============================================================================
]]

-- UTF-8 code page identifier for string conversion
local CP_UTF8 = 65001

-- COM class context: create object in-process (same process as caller)
local CLSCTX_INPROC_SERVER = 0x1

-- S_OK: Standard success return code for COM functions
local S_OK = 0

--[[
    ============================================================================
    SECTION 4: COM GUID DEFINITIONS
    ============================================================================
    GUIDs (Globally Unique Identifiers) identify COM interfaces and classes.
    These were extracted from the original C# code via decompilation.
]]

--- Helper function to create a GUID structure from its components
-- GUIDs are typically written as: XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX
-- @param data1 First 8 hex digits (32 bits)
-- @param data2 Next 4 hex digits (16 bits)
-- @param data3 Next 4 hex digits (16 bits)
-- @param d4_0 to d4_7 Final 12 hex digits as 8 bytes
-- @return GUID structure ready to pass to COM functions
local function GUID(data1, data2, data3, d4_0, d4_1, d4_2, d4_3, d4_4, d4_5, d4_6, d4_7)
    local guid = ffi.new("GUID")
    guid.Data1 = data1
    guid.Data2 = data2
    guid.Data3 = data3
    guid.Data4[0] = d4_0
    guid.Data4[1] = d4_1
    guid.Data4[2] = d4_2
    guid.Data4[3] = d4_3
    guid.Data4[4] = d4_4
    guid.Data4[5] = d4_5
    guid.Data4[6] = d4_6
    guid.Data4[7] = d4_7
    return guid
end

-- ITheme interface: Represents a Windows theme with properties like DisplayName and VisualStyle
-- GUID: D23CC733-5522-406D-8DFB-B3CF5EF52A71
local IID_ITheme = GUID(0xD23CC733, 0x5522, 0x406D, 0x8D, 0xFB, 0xB3, 0xCF, 0x5E, 0xF5, 0x2A, 0x71)

-- IThemeManager interface: Manages Windows themes (get current theme, apply new theme)
-- GUID: 0646EBBE-C1B7-4045-8FD0-FFD65D3FC792
local IID_IThemeManager = GUID(0x0646EBBE, 0xC1B7, 0x4045, 0x8F, 0xD0, 0xFF, 0xD6, 0x5D, 0x3F, 0xC7, 0x92)

-- ThemeManager COM class: The actual implementation of IThemeManager
-- GUID: C04B329E-5823-4415-9C93-BA44688947B0
local CLSID_ThemeManager = GUID(0xC04B329E, 0x5823, 0x4415, 0x9C, 0x93, 0xBA, 0x44, 0x68, 0x89, 0x47, 0xB0)

--[[
    ============================================================================
    SECTION 5: COM INTERFACE DEFINITIONS
    ============================================================================
    COM interfaces are accessed through vtables (virtual function tables).
    A vtable is an array of function pointers. Each COM object has a pointer
    to its vtable as its first member.

    Structure:
    Object -> lpVtbl -> [function pointer 0, function pointer 1, ...]

    All COM interfaces inherit from IUnknown, so the first 3 functions are
    always QueryInterface, AddRef, and Release.
]]

-- ITheme vtable: Represents a single Windows theme
ffi.cdef[[
    typedef struct IThemeVtbl {
        // === IUnknown methods (inherited by all COM interfaces) ===
        HRESULT (*QueryInterface)(void* This, const IID* riid, void** ppvObject);
        uint32_t (*AddRef)(void* This);
        uint32_t (*Release)(void* This);

        // === ITheme-specific methods ===
        // Get the display name of the theme (e.g., "Windows 10", "My Custom Theme")
        HRESULT (*get_DisplayName)(void* This, BSTR* pDisplayName);

        // Get the path to the visual style file (e.g., "C:\Windows\Resources\Themes\Aero\Aero.msstyles")
        HRESULT (*get_VisualStyle)(void* This, BSTR* pVisualStyle);
    } IThemeVtbl;

    // ITheme object structure - just a pointer to the vtable
    typedef struct ITheme {
        const IThemeVtbl* lpVtbl;
    } ITheme;
]]

-- IThemeManager vtable: Manages Windows themes
ffi.cdef[[
    typedef struct IThemeManagerVtbl {
        // === IUnknown methods (inherited by all COM interfaces) ===
        HRESULT (*QueryInterface)(void* This, const IID* riid, void** ppvObject);
        uint32_t (*AddRef)(void* This);
        uint32_t (*Release)(void* This);

        // === IThemeManager-specific methods ===
        // Get the currently active theme
        // Returns a pointer to an ITheme interface
        HRESULT (*get_CurrentTheme)(void* This, ITheme** ppTheme);

        // Apply a new theme from a .theme file
        // bstrThemePath: Full path to the .theme file
        HRESULT (*ApplyTheme)(void* This, BSTR bstrThemePath);
    } IThemeManagerVtbl;

    // IThemeManager object structure - just a pointer to the vtable
    typedef struct IThemeManager {
        const IThemeManagerVtbl* lpVtbl;
    } IThemeManager;
]]

--[[
    ============================================================================
    SECTION 6: STRING CONVERSION HELPERS
    ============================================================================
    Windows COM uses BSTR (wide character strings in UTF-16 encoding).
    Lua uses regular char* strings in UTF-8 encoding.
    These helpers convert between the two formats.
]]

--- Convert a BSTR (wide string) to a Lua string (UTF-8)
-- @param wstr BSTR pointer from a COM function
-- @return Lua string in UTF-8, or nil if conversion fails
local function wstring_to_string(wstr)
    -- Check for NULL pointer
    if wstr == nil then return nil end

    -- Get the length of the BSTR (in characters, not bytes)
    local len = oleaut32.SysStringLen(wstr)
    if len == 0 then return "" end

    -- First call: Get the required buffer size for UTF-8 conversion
    -- Passing NULL as the output buffer makes the function return the required size
    local size = ffi.C.WideCharToMultiByte(CP_UTF8, 0, wstr, len, nil, 0, nil, nil)
    if size == 0 then return nil end

    -- Allocate a buffer for the UTF-8 string (+ 1 for null terminator)
    local buf = ffi.new("char[?]", size + 1)

    -- Second call: Actually perform the conversion
    ffi.C.WideCharToMultiByte(CP_UTF8, 0, wstr, len, buf, size, nil, nil)

    -- Convert the C buffer to a Lua string
    return ffi.string(buf, size)
end

--- Convert a Lua string (UTF-8) to a BSTR (wide string)
-- @param str Lua string in UTF-8
-- @return BSTR pointer (must be freed with SysFreeString!), or nil if conversion fails
local function string_to_wstring(str)
    -- Check for nil
    if str == nil then return nil end

    -- First call: Get the required buffer size for UTF-16 conversion
    local size = ffi.C.MultiByteToWideChar(CP_UTF8, 0, str, #str, nil, 0)
    if size == 0 then return nil end

    -- Allocate a buffer for the wide string (+ 1 for null terminator)
    local wbuf = ffi.new("OLECHAR[?]", size + 1)

    -- Second call: Actually perform the conversion
    ffi.C.MultiByteToWideChar(CP_UTF8, 0, str, #str, wbuf, size)
    wbuf[size] = 0  -- Add null terminator

    -- Allocate a BSTR from the wide string buffer
    -- This creates a proper BSTR with length prefix
    return oleaut32.SysAllocString(wbuf)
end

--[[
    ============================================================================
    SECTION 7: THEMETOOL CLASS
    ============================================================================
    Main class that encapsulates all theme management functionality.
    Uses object-oriented programming in Lua (tables with metatables).
]]

local ThemeTool = {}
ThemeTool.__index = ThemeTool  -- Make ThemeTool the metatable for instances

--- Constructor: Create a new ThemeTool instance
-- Initializes COM and creates a ThemeManager COM object
-- @return ThemeTool instance
-- @throws Error if COM initialization or ThemeManager creation fails
function ThemeTool:new()
    local obj = setmetatable({}, self)

    -- Step 1: Initialize the COM library
    -- This must be called before using any COM functions
    -- nil parameter means use default settings
    local hr = ole32.CoInitialize(nil)

    -- Check for success: S_OK (0x00000000) or S_FALSE (0x00000001)
    -- S_FALSE means COM was already initialized on this thread, which is fine
    if hr ~= S_OK and hr ~= 0x00000001 then
        error("Failed to initialize COM")
    end

    -- Step 2: Create an instance of the ThemeManager COM object
    -- ppv is an output parameter that will receive the interface pointer
    local ppv = ffi.new("void*[1]")

    hr = ole32.CoCreateInstance(
        CLSID_ThemeManager,        -- The class we want to instantiate
        nil,                       -- No aggregation
        CLSCTX_INPROC_SERVER,      -- Run in the same process
        IID_IThemeManager,         -- The interface we want to use
        ppv                        -- Output: pointer to the interface
    )

    -- Check if object creation succeeded
    if hr ~= S_OK then
        ole32.CoUninitialize()  -- Clean up COM before erroring
        error(string.format("Failed to create ThemeManager instance: 0x%08X", hr))
    end

    -- Step 3: Cast the generic void* to our specific interface type
    -- This allows us to call methods through the vtable
    obj.themeManager = ffi.cast("IThemeManager*", ppv[0])

    return obj
end

--- Get the current Windows theme object
-- Internal helper method used by other methods
-- @return ITheme* pointer, or nil on failure
function ThemeTool:getCurrentTheme()
    -- Allocate an output parameter to receive the ITheme pointer
    local ppTheme = ffi.new("ITheme*[1]")

    -- Call the get_CurrentTheme method through the vtable
    -- Syntax: vtable.method(object_pointer, parameters...)
    local hr = self.themeManager.lpVtbl.get_CurrentTheme(self.themeManager, ppTheme)

    -- Check if the call succeeded
    if hr ~= S_OK then
        return nil
    end

    -- Return the ITheme pointer (first element of the output array)
    return ppTheme[0]
end

--- Get the display name of the current Windows theme
-- @return String containing the theme's display name (e.g., "Windows 10", "My Custom Theme")
function ThemeTool:getCurrentThemeName()
    -- Get the current theme object
    local theme = self:getCurrentTheme()
    if not theme then
        return ""
    end

    -- Allocate output parameter for the BSTR
    local pDisplayName = ffi.new("BSTR[1]")

    -- Call get_DisplayName through the ITheme vtable
    local hr = theme.lpVtbl.get_DisplayName(theme, pDisplayName)

    if hr ~= S_OK then
        -- Don't forget to release the theme object before returning!
        theme.lpVtbl.Release(theme)
        return ""
    end

    -- Convert BSTR to Lua string
    local name = wstring_to_string(pDisplayName[0])

    -- Free the BSTR (COM allocated it, we must free it)
    oleaut32.SysFreeString(pDisplayName[0])

    -- Release the theme object (decrement reference count)
    theme.lpVtbl.Release(theme)

    return name or ""
end

--- Get the filename of the current visual style
-- Visual styles are .msstyles files that define the appearance of Windows UI elements
-- @return String containing just the filename (e.g., "Aero.msstyles", "aero.msstyles")
function ThemeTool:getCurrentVisualStyleName()
    -- Get the current theme object
    local theme = self:getCurrentTheme()
    if not theme then
        return ""
    end

    -- Allocate output parameter for the BSTR
    local pVisualStyle = ffi.new("BSTR[1]")

    -- Get the full path to the visual style file
    local hr = theme.lpVtbl.get_VisualStyle(theme, pVisualStyle)

    if hr ~= S_OK then
        theme.lpVtbl.Release(theme)
        return ""
    end

    -- Convert BSTR to Lua string
    local path = wstring_to_string(pVisualStyle[0])

    -- Free the BSTR
    oleaut32.SysFreeString(pVisualStyle[0])

    -- Release the theme object
    theme.lpVtbl.Release(theme)

    if not path or path == "" then
        return ""
    end

    -- Extract just the filename from the full path
    -- Pattern explanation: ([^\\]+)$ matches the last sequence of non-backslash characters
    -- This gives us everything after the last backslash (the filename)
    return path:match("([^\\]+)$") or ""
end

--- Apply a new Windows theme from a .theme file
-- @param themePath Full path to a .theme file (e.g., "C:\Windows\Resources\Themes\aero.theme")
-- @return Boolean: true if theme was applied successfully, false otherwise
function ThemeTool:changeTheme(themePath)
    -- Convert Lua string to BSTR
    local bstrPath = string_to_wstring(themePath)
    if not bstrPath then
        return false
    end

    -- Call ApplyTheme through the vtable
    local hr = self.themeManager.lpVtbl.ApplyTheme(self.themeManager, bstrPath)

    -- Free the BSTR we allocated
    oleaut32.SysFreeString(bstrPath)

    -- Return success/failure
    return hr == S_OK
end

--- Check if Windows themes are currently active
-- @return String: "running" if themes are active, "stopped" if disabled
function ThemeTool:getThemeStatus()
    -- Call the UxTheme.dll API directly (doesn't need COM)
    -- Returns true if themes/visual styles are enabled, false otherwise
    return uxtheme.IsThemeActive() and "running" or "stopped"
end

--- Clean up resources and uninitialize COM
-- IMPORTANT: Must be called when done using the ThemeTool instance!
-- Failing to call this will leak COM resources
function ThemeTool:close()
    -- Release the ThemeManager COM object if we have one
    if self.themeManager then
        -- Call Release to decrement reference count
        -- When count reaches 0, the object will delete itself
        self.themeManager.lpVtbl.Release(self.themeManager)
        self.themeManager = nil
    end

    -- Uninitialize COM for this thread
    ole32.CoUninitialize()
end

--[[
    ============================================================================
    SECTION 8: MAIN FUNCTION AND COMMAND-LINE INTERFACE
    ============================================================================
    Parse command-line arguments and execute the appropriate command.
]]

--- Main entry point for the application
-- @param args Array of command-line arguments (args[1] is the command)
local function main(args)
    -- Require at least one argument (the command)
    if #args < 1 then
        return
    end

    -- Get the command and convert to lowercase for case-insensitive matching
    local command = args[1]:lower()
    local result = ""

    -- Wrap execution in pcall (protected call) to catch any errors
    -- This prevents the script from crashing and showing ugly error messages
    local success, err = pcall(function()
        -- Dispatch to the appropriate command handler
        if command == "getcurrentthemename" then
            -- Get and display the current theme's display name
            local tool = ThemeTool:new()
            result = tool:getCurrentThemeName()
            tool:close()

        elseif command == "changetheme" then
            -- Apply a new theme from a .theme file
            -- Requires a second argument: the path to the theme file
            if #args < 2 then
                return
            end
            local tool = ThemeTool:new()
            tool:changeTheme(args[2])
            tool:close()

        elseif command == "getcurrentvisualstylename" then
            -- Get and display the current visual style filename
            local tool = ThemeTool:new()
            result = tool:getCurrentVisualStyleName()
            tool:close()

        elseif command == "getthemestatus" then
            -- Check if themes are active and display status
            local tool = ThemeTool:new()
            result = tool:getThemeStatus()
            tool:close()
        end
        -- If command doesn't match any of the above, do nothing
    end)

    -- If an error occurred, suppress it and just print empty string
    -- This matches the behavior of the original C# version
    if not success then
        result = ""
    end

    -- Print the result to stdout
    print(result)
end

--[[
    ============================================================================
    SCRIPT EXECUTION
    ============================================================================
    When this script is run, the 'arg' table contains command-line arguments.
    arg[0] is the script name, arg[1] is the first argument, etc.
]]

-- Run main with command-line arguments
main(arg)
