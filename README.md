# PinToTop
"Always On Top" for Windows

## Usage
Once launched, PinToTop will stay in the tray. You can click the tray icon to select a window to stay on top.
Or you can use the hotkey "Ctrl+Alt+T" to toggle on-top for the focused window.

## Build
Visual Studio 2019 with C++ & UWP workloads and Windows 10 SDK 10.0.18362.0 is required.
```
msbuild -t:restore -p:RestorePackagesConfig=true
msbuild -t:rebuild -p:Configuration=Release
```
