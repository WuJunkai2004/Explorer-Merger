# Explorer-Merger

**Explorer-Merger** is a Windows utility designed to automatically merge newly opened File Explorer windows into tabs of an existing Explorer window. This project aims to reduce window clutter and provide a more organized file management experience on Windows 11.

This project is inspired by [ExplorerTabUtility](https://github.com/w4po/ExplorerTabUtility).

## Key Features

- **Automatic Tab Merging:** Intercepts newly created Windows Explorer (`CabinetWClass`) windows and automatically merges them into an existing Explorer window as a new tab.
- **Smart Deduplication:** If the target folder path is already open in an existing tab, the utility will focus on that tab instead of creating a duplicate one.
- **Lightweight & Efficient:** Uses Windows WinEvents and COM interfaces (`IShellWindows`, `IWebBrowser2`, `IShellBrowser`) for low-overhead window monitoring and control.
- **Dual Mode Support:**
    - **Windhawk Mod:** Can be built and loaded as a [Windhawk](https://windhawk.net/) mod for seamless background integration.
    - **Standalone Debug Tool:** Can be compiled as a standalone executable for testing and development.

## How it Works

The utility uses `SetWinEventHook` to listen for the `EVENT_OBJECT_SHOW` event on File Explorer windows. When a new Explorer window is detected, the following logic is executed:
1.  **Path Identification:** Resolves the target folder path of the new window using COM interfaces.
2.  **Window Search:** Searches for any already open Explorer window using `IShellWindows`.
3.  **Merge Logic:**
    - If an existing window is found, the new window is immediately hidden to avoid visual flicker.
    - It checks if the target path is already open in one of the existing tabs.
    - If a matching tab is found, it focuses on that tab.
    - If no matching tab is found, it sends a command to the existing window to create a new tab and navigates it to the target path.
4.  **Cleanup:** Closes the redundant new window after the merge is complete.

## Requirements

- **Windows 11:** Requires native File Explorer tab support (introduced in Windows 11 22H2).
- **C++ Compiler:** MSVC (Visual Studio) is recommended.
- **Python 3:** Required for running the build and compilation scripts.

## Installation & Building

The project uses a `Makefile` to manage the build process.
You must have Python 3 installed and added to your system PATH to use the build commands.

### 1. Building as a Windhawk Mod

To assemble the final Windhawk mod (`dist/Explorer-Merger.cpp`):

```powershell
make mod
```

### 2. Building as a Standalone Debug Executable

To compile the standalone debug tool (`dist/Explorer-Merger-Debug.exe`):

```powershell
make debug
```

> **Note:** Compiling the debug tool requires a working MSVC environment (accessible via `vcvarsall.bat`).

### 3. Cleaning Up

To remove the `dist` directory and its contents:

```powershell
make clean
```

## Credits

Special thanks to [w4po](https://github.com/w4po) for the inspiration from [ExplorerTabUtility](https://github.com/w4po/ExplorerTabUtility).

## License

This project is licensed under the [MIT License](LICENSE).
