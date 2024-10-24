#Requires AutoHotkey v2.0
#SingleInstance Force

; Load the DLL
dllPath := A_ScriptDir "\VirtualDesktopAccessor.dll"
if !FileExist(dllPath) {
    MsgBox "VirtualDesktopAccessor.dll not found in script directory!"
    ExitApp
}

; Load DLL functions
hVirtualDesktopAccessor := DllCall("LoadLibrary", "Str", dllPath, "Ptr")

GetCurrentDesktopNumber := DllCall("GetProcAddress", "Ptr", hVirtualDesktopAccessor, "AStr", "GetCurrentDesktopNumber", "Ptr")
GetDesktopCount := DllCall("GetProcAddress", "Ptr", hVirtualDesktopAccessor, "AStr", "GetDesktopCount", "Ptr")
MoveWindowToDesktopNumber := DllCall("GetProcAddress", "Ptr", hVirtualDesktopAccessor, "AStr", "MoveWindowToDesktopNumber", "Ptr")
IsPinnedWindow := DllCall("GetProcAddress", "Ptr", hVirtualDesktopAccessor, "AStr", "IsPinnedWindow", "Ptr")
PinWindow := DllCall("GetProcAddress", "Ptr", hVirtualDesktopAccessor, "AStr", "PinWindow", "Ptr")
UnPinWindow := DllCall("GetProcAddress", "Ptr", hVirtualDesktopAccessor, "AStr", "UnPinWindow", "Ptr")

; Configuration file paths
configFile := A_ScriptDir "\desktop_mappings.ini"
pinnedWindowsFile := A_ScriptDir "\pinned_windows.txt"

; Global variables for GUI elements
global mappingsGui := ""
global mappingsControls := []
global pinnedWindowsList := ""
global windowInfoText := ""
global statusBar := ""
global errorLog := ""
global autoApplyEnabled := false

; Add window event handling
SetTimer MonitorWindows, 1000  ; Check every second when auto-apply is enabled

; Window monitoring function
MonitorWindows() {
    global autoApplyEnabled
    static lastWindowList := ""
    
    if (!autoApplyEnabled) {
        return
    }
    
    ; Get current window list
    currentWindows := ""
    windowList := WinGetList()
    for hwnd in windowList {
        try {
            title := WinGetTitle(hwnd)
            if (title != "") {
                currentWindows .= title "`n"
            }
        }
    }
    
    ; If window list changed, apply rules
    if (currentWindows != lastWindowList) {
        LogError("Window list changed - auto-applying rules")
        lastWindowList := currentWindows
        ApplyCurrentMappings()
    }
}

class WindowMapping {
    __New(pattern, desktopNumber) {
        this.pattern := pattern
        this.desktopNumber := desktopNumber
    }
}

; Array join helper function
ArrayJoin(arr, delimiter) {
    result := ""
    for index, value in arr {
        if (index > 1)
            result .= delimiter
        result .= value
    }
    return result
}

; Log error message
LogError(message) {
    if errorLog {
        errorLog.Value := FormatTime(, "yyyy-MM-dd HH:mm:ss") ": " message "`n" errorLog.Value
    }
}

; Load mappings from file
LoadMappings() {
    mappings := []
    if !FileExist(configFile) {
        try {
            defaultContent := "; Window Mappings Configuration`n"
                . "; Format: window_pattern|desktop_number`n"
                . "; Use (process.exe) format to match by process name`n`n"
                . "; Example mappings`n"
                . "example project - Visual Studio Code|1`n"
                . "Outlook|2`n"
                . "(chrome.exe)|3`n`n"
                . "; Add your own mappings below"
            FileAppend defaultContent, configFile
            LogError("Created new mappings file with default content")
        } catch as err {
            LogError("Error creating mappings file: " err.Message)
        }
    }
    
    Try {
        Loop Read configFile {
            ; Skip empty lines and comments
            line := Trim(A_LoopReadLine)
            if (line = "" || SubStr(line, 1, 1) = ";") {
                continue
            }
            parts := StrSplit(line, "|")
            if (parts.Length >= 2) {
                pattern := Trim(parts[1])
                desktopNum := Integer(Trim(parts[2]))
                if (pattern != "" && desktopNum > 0) {
                    mappings.Push(WindowMapping(pattern, desktopNum))
                }
            }
        }
        LogError("Successfully loaded " mappings.Length " mappings from file")
    } catch as err {
        LogError("Error loading mappings: " err.Message)
    }
    return mappings
}

; Save mappings to file
SaveMappings(mappings) {
    try {
        ; Read existing file to preserve comments
        existingLines := []
        if FileExist(configFile) {
            Loop Read configFile {
                line := A_LoopReadLine
                if (SubStr(Trim(line), 1, 1) = ";" || Trim(line) = "") {
                    existingLines.Push(line)
                }
            }
        }
        
        ; Create new content with preserved comments
        fileContent := ""
        for line in existingLines {
            fileContent .= line "`n"
        }
        
        ; Add mappings
        if (existingLines.Length > 0) {
            fileContent .= "`n"  ; Add spacing after comments
        }
        for mapping in mappings {
            fileContent .= mapping.pattern "|" mapping.desktopNumber "`n"
        }
        
        FileDelete configFile
        FileAppend fileContent, configFile
        LogError("Successfully saved " mappings.Length " mappings to file")
        UpdateStatusBar("Mappings saved successfully")
    } catch as err {
        LogError("Error saving mappings: " err.Message)
        UpdateStatusBar("Error saving mappings")
    }
}

; Load pinned windows list
LoadPinnedWindows() {
    pinnedWindows := []
    if !FileExist(pinnedWindowsFile) {
        try {
            defaultContent := "; Pinned Windows Configuration`n"
                . "; One pattern per line`n"
                . "; Use (process.exe) format to match by process name`n`n"
                . "; Example pinned applications`n"
                . "Virtual Desktop Manager`n"
                . "Microsoft Teams`n`n"
                . "; Add your own pinned windows below"
            FileAppend defaultContent, pinnedWindowsFile
            LogError("Created new pinned windows file with default content")
        } catch as err {
            LogError("Error creating pinned windows file: " err.Message)
        }
    }
    
    Try {
        Loop Read pinnedWindowsFile {
            ; Skip empty lines and comments
            line := Trim(A_LoopReadLine)
            if (line = "" || SubStr(line, 1, 1) = ";") {
                continue
            }
            if (line != "") {
                pinnedWindows.Push(line)
                LogError("Added pinned window pattern: '" line "'")
            }
        }
        LogError("Successfully loaded " pinnedWindows.Length " pinned window patterns")
    } catch as err {
        LogError("Error loading pinned windows: " err.Message)
    }
    return pinnedWindows
}

; Save pinned windows list
SavePinnedWindows(pinnedWindows) {
    try {
        ; Read existing file to preserve comments
        existingLines := []
        if FileExist(pinnedWindowsFile) {
            Loop Read pinnedWindowsFile {
                line := A_LoopReadLine
                if (SubStr(Trim(line), 1, 1) = ";" || Trim(line) = "") {
                    existingLines.Push(line)
                }
            }
        }
        
        ; Create new content with preserved comments
        content := ""
        for line in existingLines {
            content .= line "`n"
        }
        
        ; Add pinned windows
        if (existingLines.Length > 0) {
            content .= "`n"  ; Add spacing after comments
        }
        for window in pinnedWindows {
            if (window != "") {
                content .= Trim(window) "`n"
            }
        }
        
        FileDelete pinnedWindowsFile
        FileAppend content, pinnedWindowsFile
        LogError("Successfully saved " pinnedWindows.Length " pinned window patterns")
        UpdateStatusBar("Pinned windows list saved")
    } catch as err {
        LogError("Error saving pinned windows: " err.Message)
        UpdateStatusBar("Error saving pinned windows")
    }
}

; Get current window information with grouped windows
GetWindowInfo() {
    info := "Pinned Windows:`n-------------------`n"
    windowList := WinGetList()
    
    ; Separate windows into pinned and unpinned
    pinnedWindows := []
    unpinnedWindows := []
    
    for hwnd in windowList {
        try {
            title := WinGetTitle(hwnd)
            if (title != "") {
                process := WinGetProcessName(hwnd)
                isPinned := DllCall(IsPinnedWindow, "Ptr", hwnd)
                windowInfo := title " (" process ")"
                
                if (isPinned) {
                    pinnedWindows.Push(windowInfo)
                } else {
                    unpinnedWindows.Push(windowInfo)
                }
            }
        }
    }
    
    ; Add pinned windows to info
    for window in pinnedWindows {
        info .= window "`n"
    }
    
    info .= "`nUnpinned Windows:`n-------------------`n"
    
    ; Add unpinned windows to info
    for window in unpinnedWindows {
        info .= window "`n"
    }
    
    info .= "`nVirtual Desktops:`n-------------------`n"
    desktopCount := DllCall(GetDesktopCount)
    Loop desktopCount {
        info .= "Desktop " A_Index "`n"
    }
    
    return info
}

UpdateStatusBar(message := "") {
    if (message = "") {
        currentDesktop := DllCall(GetCurrentDesktopNumber) + 1
        totalDesktops := DllCall(GetDesktopCount)
        message := "Current Desktop: " currentDesktop "/" totalDesktops
        if (autoApplyEnabled)
            message .= " | Auto-Apply: Active (checking every 1s)"
        else
            message .= " | Auto-Apply: Disabled"
    }
    statusBar.Text := message
}

; Modified apply current mappings to use the new function
ApplyCurrentMappings(*) {
    try {
        currentMappings := GetCurrentMappings()
        currentPinned := GetCurrentPinnedWindows()
        LogError("Applying mappings with " currentPinned.Length " pinned window patterns")
        ApplyMappings(currentMappings, currentPinned)
        RefreshWindowInfo()
        UpdateStatusBar("Mappings applied")
    } catch as err {
        LogError("Error applying mappings: " err.Message)
        UpdateStatusBar("Error applying mappings")
    }
}

; Simplified save function
SaveCurrentMappings(*) {
    try {
        mappings := GetCurrentMappings()
        SaveMappings(mappings)
        currentPinned := StrSplit(pinnedWindowsList.Value, "`n")
        SavePinnedWindows(currentPinned)
        UpdateStatusBar("Settings saved")
    } catch as err {
        LogError("Error saving: " err.Message)
        UpdateStatusBar("Error saving settings")
    }
}

; Handler for refreshing window info
RefreshWindowInfo(*) {
    try {
        windowInfoText.Value := GetWindowInfo()
        UpdateStatusBar("Window info refreshed")
    } catch as err {
        LogError("Error refreshing window info: " err.Message)
        UpdateStatusBar("Error refreshing window info")
    }
}

; Create mapping row controls
CreateMappingControls(gui, mapping, index, yOffset) {
    controls := Map()
    
    controls["pattern"] := gui.Add("Edit", "x10 y" (35 + yOffset) " w200", mapping.pattern)
    controls["desktop"] := gui.Add("Edit", "x220 y" (35 + yOffset) " w200 Number", mapping.desktopNumber)
    deleteBtn := gui.Add("Button", "x430 y" (35 + yOffset) " w30", "X")
    controls["delete"] := deleteBtn
    
    ; Pass the pattern as a unique identifier instead of the index
    deleteBtn.OnEvent("Click", DeleteMapping.Bind(mapping.pattern))
    
    return controls
}

; Delete mapping handler - simplified
DeleteMapping(pattern, *) {
    try {
        ; Read current mappings from file
        mappings := LoadMappings()
        
        ; Find and remove the mapping with this pattern
        for index, mapping in mappings {
            if (mapping.pattern = pattern) {
                mappings.RemoveAt(index)
                break
            }
        }
        
        ; Save updated mappings to file
        SaveMappings(mappings)
        
        ; Reload the entire GUI
        Reload()
        
    } catch as err {
        MsgBox "Error deleting mapping: " err.Message
    }
}

; Add new mapping handler
AddNewMapping(*) {
    global mappingsControls
    try {
        ; Instead of adding directly, we'll recreate the GUI with a new mapping
        mappingsControls.Push(WindowMapping("", 1))  ; Add new mapping to array
        ShowMappingsGui()  ; Recreate entire GUI with new mapping
        UpdateStatusBar("New mapping added")
    } catch as err {
        LogError("Error adding new mapping: " err.Message)
        UpdateStatusBar("Error adding mapping")
    }
}

; Modified ToggleAutoApply handler
ToggleAutoApply(*) {
    global autoApplyEnabled
    autoApplyEnabled := !autoApplyEnabled
    if (autoApplyEnabled) {
        LogError("Auto-apply enabled")
        ; Apply rules immediately when enabled
        ApplyCurrentMappings()
    } else {
        LogError("Auto-apply disabled")
    }
    UpdateStatusBar()
}

; Apply mappings and pinning rules
ApplyMappings(mappings, pinnedWindows) {
    windowList := WinGetList()
    
    ; First pass: Handle pinning
    for hwnd in windowList {
        try {
            title := WinGetTitle(hwnd)
            process := WinGetProcessName(hwnd)
            
            if (title != "") {
                ; Skip System Windows that shouldn't be unpinned
                if (title = "Program Manager" || InStr(process, "explorer.exe")) {
                    continue
                }
                
                ; Check pinning rules against both title and process name
                shouldBePinned := false
                for pattern in pinnedWindows {
                    pattern := Trim(pattern)
                    if (pattern = "") {
                        continue
                    }
                    
                    ; Check if pattern is a process pattern (enclosed in parentheses)
                    if (SubStr(pattern, 1, 1) = "(" && SubStr(pattern, -1) = ")") {
                        processPattern := SubStr(pattern, 2, StrLen(pattern) - 2)
                        if (InStr(process, processPattern)) {
                            LogError("Window '" title "' matches process pattern '" processPattern "' - will pin")
                            shouldBePinned := true
                            break
                        }
                    }
                    ; Otherwise check title and process name
                    else if (InStr(title, pattern) || InStr(process, pattern)) {
                        LogError("Window '" title "' (" process ") matches pin pattern '" pattern "' - will pin")
                        shouldBePinned := true
                        break
                    }
                }
                
                ; Apply pinning rules
                isPinned := DllCall(IsPinnedWindow, "Ptr", hwnd)
                if (shouldBePinned && !isPinned) {
                    LogError("Pinning window: " title)
                    DllCall(PinWindow, "Ptr", hwnd)
                } else if (!shouldBePinned && isPinned && !InStr(process, "explorer.exe")) {
                    LogError("Unpinning window: " title)
                    DllCall(UnPinWindow, "Ptr", hwnd)
                }
            }
        } catch as err {
            LogError("Error processing window for pinning: " err.Message)
        }
    }
    
    ; Second pass: Handle desktop mappings
    for hwnd in windowList {
        try {
            title := WinGetTitle(hwnd)
            process := WinGetProcessName(hwnd)
            
            if (title != "") {
                ; Check desktop mappings
                for mapping in mappings {
                    pattern := mapping.pattern
                    
                    ; Check if pattern is a process pattern (enclosed in parentheses)
                    if (SubStr(pattern, 1, 1) = "(" && SubStr(pattern, -1) = ")") {
                        processPattern := SubStr(pattern, 2, StrLen(pattern) - 2)
                        if (InStr(process, processPattern)) {
                            LogError("Moving window '" title "' to desktop " mapping.desktopNumber " (process match)")
                            DllCall(MoveWindowToDesktopNumber, "Ptr", hwnd, "Int", mapping.desktopNumber - 1)
                            break
                        }
                    }
                    ; Otherwise check title
                    else if (InStr(title, pattern)) {
                        LogError("Moving window '" title "' to desktop " mapping.desktopNumber " (title match)")
                        DllCall(MoveWindowToDesktopNumber, "Ptr", hwnd, "Int", mapping.desktopNumber - 1)
                        break
                    }
                }
            }
        } catch as err {
            LogError("Error processing window for desktop mapping: " err.Message)
        }
    }
}

; Create and show main GUI
ShowMappingsGui() {
    global mappingsGui, mappingsText, pinnedWindowsList, windowInfoText, statusBar, errorLog, autoApplyEnabled
    
    ; Load both mappings and pinned windows
    mappings := LoadMappings()
    pinnedWindows := LoadPinnedWindows()
    
    ; Convert mappings to text format
    mappingsString := ""
    for mapping in mappings {
        mappingsString .= mapping.pattern "|" mapping.desktopNumber "`n"
    }
    
    ; Convert pinned windows to text format
    pinnedString := ""
    for window in pinnedWindows {
        pinnedString .= window "`n"
    }
    
    mappingsGui := Gui("+Resize", "Virtual Desktop Manager")
    
    ; Mappings text area
    mappingsGui.Add("Text", "x10 y10", "Window Mappings (format: window_pattern|desktop_number):")
    mappingsText := mappingsGui.Add("Edit", "x10 y30 w450 h200 vMappings", mappingsString)
    
    ; Buttons
    applyButton := mappingsGui.Add("Button", "x10 y240", "Apply Mappings")
    applyButton.OnEvent("Click", ApplyCurrentMappings)
    
    saveButton := mappingsGui.Add("Button", "x100 y240", "Save")
    saveButton.OnEvent("Click", SaveCurrentMappings)
    
    reloadButton := mappingsGui.Add("Button", "x190 y240", "Reload Script")
    reloadButton.OnEvent("Click", (*) => Reload())
    
    autoApplyCheckbox := mappingsGui.Add("Checkbox", "x280 y243 vAutoApply", "Auto-Apply")
    autoApplyCheckbox.OnEvent("Click", ToggleAutoApply)
    ; Set checkbox state based on current auto-apply setting
    if (autoApplyEnabled) {
        autoApplyCheckbox.Value := 1
    }
    
    ; Pinned windows section
    mappingsGui.Add("Text", "x10 y280", "Pinned Windows Patterns (one per line):")
    pinnedWindowsList := mappingsGui.Add("Edit", "x10 y300 w450 h100 vPinnedWindows", pinnedString)

    mappingsGui.Add("Text", "x10 y410", "Current Windows and Desktops:")
    windowInfoText := mappingsGui.Add("Edit", "x10 y430 w450 h200 ReadOnly vWindowInfo")
    
    mappingsGui.Add("Text", "x10 y640", "Log:")
    errorLog := mappingsGui.Add("Edit", "x10 y660 w450 h100 ReadOnly vErrorLog")
    
    ; Status Bar
    statusBar := mappingsGui.Add("Text", "x10 y770 w450 h20")
    
    ; Refresh button
    refreshButton := mappingsGui.Add("Button", "x370 y410", "Refresh Window Info")
    refreshButton.OnEvent("Click", RefreshWindowInfo)
    
    ; GUI Events
    mappingsGui.OnEvent("Size", GuiResize)
    mappingsGui.OnEvent("Close", (*) => ExitApp())
    
    mappingsGui.Show()
    UpdateStatusBar()
    RefreshWindowInfo()
}

; Handle GUI resize
GuiResize(gui, minMax := false, width := false, height := false) {
    if (minMax = -1) {  ; Window minimized
        return
    }
    
    try {
        width := gui.WindowWidth
    }
    
    if (width) {
        ; Adjust control sizes
        for controls in mappingsControls {
            controls["pattern"].Move(,, width - 250)
            controls["desktop"].Move(width - 230,, 180)
            controls["delete"].Move(width - 40)
        }
        
        ; Adjust other controls
        if (pinnedWindowsList) {
            pinnedWindowsList.Move(,, width - 20)
        }
        if (windowInfoText) {
            windowInfoText.Move(,, width - 20)
        }
        if (errorLog) {
            errorLog.Move(,, width - 20)
        }
        if (statusBar) {
            statusBar.Move(,, width - 20)
        }
    }
}

; Modified to better handle pinned windows parsing
GetCurrentPinnedWindows() {
    if !pinnedWindowsList {
        return []
    }
    
    pinnedWindows := []
    lines := StrSplit(pinnedWindowsList.Value, "`n")
    for line in lines {
        line := Trim(line)
        if (line != "") {
            pinnedWindows.Push(line)
            LogError("Added pinned window pattern: '" line "'")
        }
    }
    return pinnedWindows
}

; Modified GetCurrentMappings to parse from text
GetCurrentMappings() {
    mappings := []
    
    if !mappingsText {
        return mappings
    }
    
    ; Split into lines and parse each line
    lines := StrSplit(mappingsText.Value, "`n")
    for line in lines {
        if (line = "") {
            continue
        }
        
        try {
            parts := StrSplit(line, "|")
            if (parts.Length >= 2) {
                pattern := Trim(parts[1])
                desktopNum := Integer(Trim(parts[2]))
                if (pattern != "" && desktopNum > 0) {
                    mappings.Push(WindowMapping(pattern, desktopNum))
                }
            }
        } catch as err {
            LogError("Error parsing line '" line "': " err.Message)
        }
    }
    return mappings
}

; Start the application
ShowMappingsGui()