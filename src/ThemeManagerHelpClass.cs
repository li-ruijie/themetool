/*
 * ThemeManagerHelpClass.cs
 *
 * Purpose:
 * This file provides a C# wrapper around the undocumented Windows Theme Manager COM interfaces.
 * It enables programmatic access to Windows theme settings including:
 * - Getting the current theme name
 * - Getting the current visual style name
 * - Changing the active theme
 * - Checking if themes are active
 *
 * Architecture:
 * The file contains COM interface definitions (ITheme, IThemeManager) that map to
 * undocumented Windows COM objects in uxtheme.dll. These interfaces use Platform
 * Invoke (P/Invoke) to call native Windows functions.
 *
 * Type: ThemeApi.ThemeManagerHelpClass
 * Assembly: ThemeTool, Version=0.0.0.0, Culture=neutral, PublicKeyToken=null
 */

using System;
using System.Globalization;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Security.Permissions;

#nullable disable
namespace ThemeApi;

/// <summary>
/// ThemeManagerHelpClass provides static methods for interacting with Windows themes.
/// This class acts as a facade over the COM-based IThemeManager interface.
/// </summary>
public static class ThemeManagerHelpClass
{
  // Static instance of the IThemeManager COM object
  // This is created once and reused for all theme operations
  private static ThemeManagerHelpClass.IThemeManager themeManager = (ThemeManagerHelpClass.IThemeManager) new ThemeManagerHelpClass.ThemeManagerClass();

  /// <summary>
  /// Gets the display name of the currently active Windows theme.
  /// </summary>
  /// <returns>The theme display name (e.g., "Windows 10", "Custom Theme")</returns>
  [PermissionSet(SecurityAction.LinkDemand)]
  public static string GetCurrentThemeName()
  {
    // Access the CurrentTheme property and retrieve its DisplayName
    return ThemeManagerHelpClass.themeManager.CurrentTheme.DisplayName;
  }

  /// <summary>
  /// Applies a Windows theme from a .theme file.
  /// </summary>
  /// <param name="themeFilePath">Full path to the .theme file to apply</param>
  [PermissionSet(SecurityAction.LinkDemand)]
  public static void ChangeTheme(string themeFilePath)
  {
    // Call the ApplyTheme method on the IThemeManager COM object
    ThemeManagerHelpClass.themeManager.ApplyTheme(themeFilePath);
  }

  /// <summary>
  /// Gets the filename of the currently active visual style.
  /// </summary>
  /// <returns>Visual style filename (e.g., "Aero.msstyles")</returns>
  [PermissionSet(SecurityAction.LinkDemand)]
  public static string GetCurrentVisualStyleName()
  {
    // Get the full path to the visual style file and extract just the filename
    return Path.GetFileName(ThemeManagerHelpClass.themeManager.CurrentTheme.VisualStyle);
  }

  /// <summary>
  /// Checks whether Windows themes are currently active.
  /// </summary>
  /// <returns>"running" if themes are active, "stopped" otherwise</returns>
  public static string GetThemeStatus()
  {
    // Call the native IsThemeActive function and return appropriate status string
    return !ThemeManagerHelpClass.NativeMethods.IsThemeActive() ? "stopped" : "running";
  }

  /// <summary>
  /// Main entry point for the ThemeTool command-line application.
  /// Parses command-line arguments and executes the appropriate theme operation.
  /// </summary>
  /// <param name="args">Command-line arguments</param>
  /// <remarks>
  /// Supported commands:
  /// - getcurrentthemename: Print the current theme name
  /// - getcurrentvisualstylename: Print the current visual style filename
  /// - getthemestatus: Print "running" or "stopped"
  /// - changetheme [path]: Apply a theme from the specified path
  ///
  /// If no arguments provided or an invalid command is given, the program exits silently.
  /// </remarks>
  [STAThread]  // Single-Threaded Apartment required for COM interop
  [PermissionSet(SecurityAction.LinkDemand)]
  public static void Main(string[] args)
  {
    // Exit silently if no command provided
    if (args.Length < 1)
      return;

    // Initialize output string
    string format = "";

    // Convert command to lowercase for case-insensitive comparison
    string lower = args[0].ToLower(CultureInfo.InvariantCulture);

    try
    {
      // Parse and execute the requested command

      // Command: getcurrentthemename
      if (string.Compare(lower, "getcurrentthemename") == 0)
        format = ThemeManagerHelpClass.GetCurrentThemeName();

      // Command: changetheme <path>
      else if (string.Compare(lower, "changetheme") == 0)
      {
        // Require a second argument (the theme file path)
        if (args.Length < 2)
          return;
        ThemeManagerHelpClass.ChangeTheme(args[1]);
      }

      // Command: getcurrentvisualstylename
      else if (string.Compare(lower, "getcurrentvisualstylename") == 0)
      {
        format = ThemeManagerHelpClass.GetCurrentVisualStyleName();
      }

      // Command: getthemestatus
      else
      {
        // Exit silently for unrecognized commands
        if (string.Compare(lower, "getthemestatus") != 0)
          return;
        format = ThemeManagerHelpClass.GetThemeStatus();
      }
    }
    catch
    {
      // On any error, output empty string (suppress error messages)
      format = "";
    }

    // Output the result to console
    Console.WriteLine(string.Format((IFormatProvider) CultureInfo.InvariantCulture, format));
  }

  // ==================== COM INTERFACE DEFINITIONS ====================
  // The following interfaces map to undocumented Windows COM objects
  // that manage theme operations. These GUIDs were extracted from the
  // Windows Theme Manager implementation.

  /// <summary>
  /// ITheme interface represents a single Windows theme.
  /// Provides access to theme properties like display name and visual style path.
  /// </summary>
  /// <remarks>
  /// GUID: D23CC733-5522-406D-8DFB-B3CF5EF52A71
  /// This is an undocumented Windows interface.
  /// </remarks>
  [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
  [Guid("D23CC733-5522-406D-8DFB-B3CF5EF52A71")]
  [ComImport]
  public interface ITheme
  {
    /// <summary>
    /// Gets the human-readable display name of the theme.
    /// </summary>
    /// <returns>Theme display name (e.g., "Windows 10", "Unsaved Theme")</returns>
    [DispId(1610678272 /*0x60010000*/)]
    string DisplayName { [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)] [return: MarshalAs(UnmanagedType.BStr)] get; }

    /// <summary>
    /// Gets the full path to the visual style file (.msstyles).
    /// </summary>
    /// <returns>Full path to .msstyles file (e.g., "C:\Windows\Resources\Themes\Aero\Aero.msstyles")</returns>
    [DispId(1610678273 /*0x60010001*/)]
    string VisualStyle { [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)] [return: MarshalAs(UnmanagedType.BStr)] get; }
  }

  /// <summary>
  /// IThemeManager interface manages Windows themes.
  /// Provides methods to get the current theme and apply new themes.
  /// </summary>
  /// <remarks>
  /// GUID: 0646EBBE-C1B7-4045-8FD0-FFD65D3FC792
  /// This is an undocumented Windows interface.
  /// </remarks>
  [Guid("0646EBBE-C1B7-4045-8FD0-FFD65D3FC792")]
  [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
  [ComImport]
  public interface IThemeManager
  {
    /// <summary>
    /// Gets the currently active Windows theme.
    /// </summary>
    /// <returns>ITheme object representing the current theme</returns>
    [DispId(1610678272 /*0x60010000*/)]
    ThemeManagerHelpClass.ITheme CurrentTheme { [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    /// <summary>
    /// Applies a Windows theme from a .theme file.
    /// </summary>
    /// <param name="bstrThemePath">Full path to the .theme file to apply</param>
    [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
    void ApplyTheme([MarshalAs(UnmanagedType.BStr), In] string bstrThemePath);
  }

  /// <summary>
  /// ThemeManager interface wrapper for IThemeManager.
  /// This is used for COM class instantiation.
  /// </summary>
  /// <remarks>
  /// GUID: A2C56C2A-E63A-433E-9953-92E94F0122EA
  /// </remarks>
  [CoClass(typeof (ThemeManagerHelpClass.ThemeManagerClass))]
  [Guid("A2C56C2A-E63A-433E-9953-92E94F0122EA")]
  [ComImport]
  public interface ThemeManager : ThemeManagerHelpClass.IThemeManager
  {
  }

  /// <summary>
  /// ThemeManagerClass is the concrete implementation of IThemeManager.
  /// This class is instantiated via COM to access Windows theme management.
  /// </summary>
  /// <remarks>
  /// CLSID: C04B329E-5823-4415-9C93-BA44688947B0
  /// This is the COM class ID used to create instances of the Theme Manager.
  /// </remarks>
  [ClassInterface(ClassInterfaceType.None)]
  [TypeLibType(TypeLibTypeFlags.FCanCreate)]
  [Guid("C04B329E-5823-4415-9C93-BA44688947B0")]
  [ComImport]
  public class ThemeManagerClass :
    ThemeManagerHelpClass.ThemeManager,
    ThemeManagerHelpClass.IThemeManager
  {
    /// <summary>
    /// Applies a theme from the specified path.
    /// </summary>
    /// <param name="bstrThemePath">Path to .theme file</param>
    [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
    public virtual extern void ApplyTheme([MarshalAs(UnmanagedType.BStr), In] string bstrThemePath);

    /// <summary>
    /// Gets the currently active theme.
    /// </summary>
    [DispId(1610678272 /*0x60010000*/)]
    public virtual extern ThemeManagerHelpClass.ITheme CurrentTheme { [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    /// <summary>
    /// COM constructor - instantiates the Windows Theme Manager COM object.
    /// </summary>
    [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
    public extern ThemeManagerClass();
  }

  /// <summary>
  /// NativeMethods contains P/Invoke declarations for native Windows API functions.
  /// This class provides direct access to unmanaged functions in Windows DLLs.
  /// </summary>
  private static class NativeMethods
  {
    /// <summary>
    /// Checks whether visual themes are active on the system.
    /// </summary>
    /// <returns>True if themes are active, false otherwise</returns>
    /// <remarks>
    /// This function is exported from UxTheme.dll and is part of the
    /// documented Windows Theme API.
    /// </remarks>
    [DllImport("UxTheme.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool IsThemeActive();
  }
}
