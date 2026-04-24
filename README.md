# PortableApps Launcher Maker

PortableApps Launcher Maker is a Windows Tkinter desktop app for building **PortableApps.com-style project folders** from a normal application EXE.

It helps generate the pieces needed for a portable app package, including:
- `appinfo.ini`
- `launcher.ini`
- optional `installer.ini`
- PortableApps icon sizes
- `help.html` and help assets
- splash assets
- preview output before build

It can also call the **PortableApps.com Launcher Generator** to build the final portable launcher EXE when that tool is available.

## Features

- creates the standard `App`, `Data`, and `Other` structure
- copies the selected application's folder
- extracts embedded icons from the selected EXE
- supports registry, icon, splash, and installer settings
- previews generated folder layout and INI output live
- supports both 64-bit and 32-bit packaged app builds

## Running From Source

```powershell
python -m app.portableapps_main
```

Or install in editable mode:

```powershell
pip install -e .
portableapps-launcher-maker
```

## Release Builds

The packaged releases use **PyInstaller one-folder builds**. Keep the whole output folder together; do not copy only the `.exe`.

### 64-bit

Requires a **64-bit Python** with PyInstaller available in that interpreter.

```powershell
powershell -ExecutionPolicy Bypass -File .\build_portableapps_release.ps1 -Python64 "C:\Path\To\Python64\python.exe"
```

Output:
- `dist\PortableAppsLauncherMaker\PortableAppsLauncherMaker.exe`

### 32-bit / x86

Requires:
- a **32-bit Python**
- `PyInstaller` installed into that 32-bit Python

```powershell
powershell -ExecutionPolicy Bypass -File .\build_portableapps_release_x86.ps1 -Python32 "C:\Path\To\Python32\python.exe"
```

You can also set `PORTABLEAPPS_X86_PYTHON` first and run the script without arguments.

Output:
- `dist\PortableAppsLauncherMaker-x86\PortableAppsLauncherMaker-x86.exe`

## Downloads

For GitHub, the best distribution path is:
- commit **source code** to the repo
- upload packaged builds to **GitHub Releases**

Recommended release assets:
- `PortableAppsLauncherMaker-win64.zip`
- `PortableAppsLauncherMaker-win32.zip`

## Repository Layout

- `app/`: application source
- `tests/`: automated tests
- `build_portableapps_release.ps1`: 64-bit release build
- `build_portableapps_release_x86.ps1`: 32-bit release build

## Notes

- `build/` and `dist/` are generated output and should not normally be committed
- generated PyInstaller `.spec` files are intentionally ignored
- building the final portable launcher EXE expects the **PortableApps.com Launcher Generator** to be installed and available
