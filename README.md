# ThemeTool - LuaJIT Port

The themetool.exe is an tool from Windows 7/10 that allows users to manage Windows themes programmatically. This is a LuaJIT port of the original C# implementation, utilizing LuaJIT's FFI to call Windows APIs and COM interfaces.

## Requirements

- LuaJIT (tested with LuaJIT 2.0+)
- Windows OS (uses Windows-specific APIs)

## Usage

The tool supports the following commands:

### Get Theme Status
Check if Windows themes are currently active:
```cmd
luajit themetool.lua getthemestatus
```
Output: `running` or `stopped`

### Get Current Theme Name
Get the display name of the current Windows theme:
```cmd
luajit themetool.lua getcurrentthemename
```

### Get Current Visual Style Name
Get the filename of the current visual style:
```cmd
luajit themetool.lua getcurrentvisualstylename
```

### Change Theme
Apply a new Windows theme from a .theme file:
```cmd
luajit themetool.lua changetheme "C:\Path\To\Theme.theme"
```

## Implementation Details

This port uses LuaJIT's FFI to:
- Call the Windows UxTheme.dll API (`IsThemeActive`)
- Interact with COM interfaces (`IThemeManager`, `ITheme`)
- Manage Windows themes programmatically

### COM Interfaces Used

- **CLSID_ThemeManager**: `C04B329E-5823-4415-9C93-BA44688947B0`
- **IID_IThemeManager**: `0646EBBE-C1B7-4045-8FD0-FFD65D3FC792`
- **IID_ITheme**: `D23CC733-5522-406D-8DFB-B3CF5EF52A71`

## Differences from C# Version

1. Direct FFI calls instead of P/Invoke
2. Manual COM interface vtable management
3. Explicit memory management for BSTR strings
4. No .NET security permissions required

## License

Ported from the original C# implementation. Use at your own risk.
