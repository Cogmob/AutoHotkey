# Virtual Desktop Manager for Windows

A personal AutoHotkey solution for managing Windows virtual desktops that I built to solve my own needs. Feel free to fork it and modify it to suit your requirements!

## 🔍 What Is This?

This is an AutoHotkey script that helps organize windows across virtual desktops by:
- Moving specific windows to designated desktops
- Pinning applications to appear on all desktops
- Auto-applying rules as new windows appear

## ⚡ Quick Start

1. Install [AutoHotkey v2.0](https://www.autohotkey.com/) or later
2. Clone/download this repository
3. Copy `desktop_mappings.ini.example` to `desktop_mappings.ini`
4. Copy `pinned_windows.txt.example` to `pinned_windows.txt`
5. Run `VirtualDesktopManager.ahk`
6. Configure your rules in the GUI

## 📝 Configuration Examples

### Window Mappings (desktop_mappings.ini)
```ini
; Format: window_pattern|desktop_number
example project - Visual Studio Code|1
Outlook|2
(chrome.exe)|3  ; Use parentheses for process name matching
```

### Pinned Windows (pinned_windows.txt)
```
Virtual Desktop Manager
Microsoft Teams
(OUTLOOK.EXE)  ; Use parentheses for process name matching
```

## 🔄 Forking Guide

1. Fork the repository
2. Update ATTRIBUTION.md with your fork's information
3. Modify the code to suit your needs
4. Update the documentation for your changes
5. If you publish your fork, maintain the attribution chain

## 🛠️ Modification Guide

This script is designed to be forked and modified. Here's where to start:

- Window management logic is in `ApplyMappings()` function
- GUI code is in `ShowMappingsGui()`
- Auto-apply functionality is in `MonitorWindows()`
- Configuration handling is in `LoadMappings()` and `LoadPinnedWindows()`

## ⚠️ Important Notes

- This is a personal tool shared as-is
- No active maintenance or support provided
- **Please fork and modify** rather than expecting updates or fixes
- If you make improvements, consider sharing them with others by publishing your fork!

## 📋 Requirements

- Windows 10/11
- AutoHotkey v2.0+
- VirtualDesktopAccessor.dll (included)

## ⚖️ License

MIT License - see LICENSE file. 

In short:
- Use it however you want
- Modify it however you want
- Keep the attribution chain
- No warranty provided

## 👋 Attribution

This is part of a chain of Virtual Desktop Manager implementations. See ATTRIBUTION.md for the full history.

Originally created by Luke Avery. Uses [VirtualDesktopAccessor.dll](https://github.com/Ciantic/VirtualDesktopAccessor) for virtual desktop interaction.

If you find this useful and improve upon it, please maintain the attribution chain in ATTRIBUTION.md when creating your fork. Enjoy!