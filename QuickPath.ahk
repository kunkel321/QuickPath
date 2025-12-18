#Requires AutoHotkey v2.0
#SingleInstance Force

/*
===============================================================================
Title:        QuickPath
Version:      12-17-2025 
Made by:      kunkel321
AHK forum:    https://www.autohotkey.com/boards/viewtopic.php?f=83&t=134987
GitHub repo:  https://github.com/kunkel321/QuickPath
QuickPath has a minimal interface.  When a Windows 'Open' or 'Save as' dialog is opened, whatever (if any) folders are open in XYplorer, DirectoryOpus, or Windows 10 Explorer, will be listed in a popup menu.  Click a menu item to quickly add that folder to the Windows dialog. Active DOpus or Xyplorer tabs are marked with a blue arrow icon.  QuickPath was inspired by the AHK v1 app, QuickSwitch by NotNull, which has the  same functionality. https://www.voidtools.com/forum/viewtopic.php?f=2&t=9881 Some of the code came from the AHK v2 app, QuickerAccess by william_ahk.  https://www.autohotkey.com/boards/viewtopic.php?f=83&t=134379&sid The XYplorer parts are based on FileManRedirect by WKen https://www.autohotkey.com/boards/viewtopic.php?f=83&t=135693  Claude AI was used extensively, though hours of human input was needed.  An additional important update from WKen made the cmd window flicker go away. :)   QuickPath has not been tested on Win 11.
;===============================================================================
*/

class QuickPath {
    static UserHotKey := "!q" ; Set custom hotkey here (only)
    static UserDopusPath := "\GPSoftware\Directory Opus\" ; DOpus path with \backSlashes\
    static hActvWnd := ""  ; Handle to previously active window
    static Enabled := true  ; Flag to control whether the menu should appear
    static dialogCheckTimer := ""
    static lastCheckTime := 0  ; Time of last folder check
    static checkTimeout := 2000  ; Time to wait before checking again (2 seconds)
    static XYplorerMaxTabLoops := 5  ; Max number of tabs to check per pane in XYplorer (recommend setting to 2-5)
    static XYplorerClipWaitTimeout := 0.25  ; ClipWait timeout in seconds (0.2-0.5 recommended, too low = missing paths)
    static pathsCache := []  ; Cache for found paths
    static activeTabsCache := Map()  ; New: Store which paths are in active tabs
    static lastDialogHwnd := 0  ; Last dialog window we checked
    static hotkeyBound := false

    ; Initialization
    static __New() {
        ; Monitor for #32770 class windows (file dialogs)
        SetTimer () => this.CheckForFileDialog(), 100
    }

    ; Check for Open/Save dialogs and handle them
    static CheckForFileDialog() {
        dialogActive := false
        if WinActive("ahk_class #32770") {
            hwnd := WinExist("A")
            dialogActive := this.IsFileDialog(hwnd)
            if dialogActive {
                if !this.hotkeyBound { ; Ensure hotkey is bound when dialog is active
                    Hotkey this.UserHotKey, (*) => this.HotkeyHandler(), "On"
                    this.hotkeyBound := true
                }
                if (hwnd = this.lastDialogHwnd) ; Skip if we've already checked this dialog
                    return

                this.lastDialogHwnd := hwnd
                this.hActvWnd := hwnd  ; Store the dialog handle
                if this.Enabled {  ; Check paths and show menu
                    this.CheckAndShowPaths()
                }
                this.StartDialogMonitor() ; Start monitoring this dialog
            }
        }
        if !dialogActive && this.hotkeyBound { ; Disable hotkey if no dialog is active
            Hotkey this.UserHotKey, "Off"
            this.hotkeyBound := false
        }
    }

    static HotkeyHandler() {  ; Handle Alt+Q hotkey
        if WinActive("ahk_id " this.hActvWnd) {
            this.Enabled := true  ; Re-enable the menu
            this.CheckAndShowPaths()  ; Show the menu immediately
        }
    }

    static CheckAndShowPaths() { ; Check paths and show menu if appropriate
        ; Only check for paths if enough time has passed
        if (A_TickCount - this.lastCheckTime > this.checkTimeout) {
            this.pathsCache := this.GetOpenPaths()
            this.lastCheckTime := A_TickCount
        } 
        if this.pathsCache.Length ; Show menu if we have paths
            this.ShowPathsMenu()
    }

    static StartDialogMonitor() { ; Start monitoring the current dialog
        if this.dialogCheckTimer  ; Clear any existing timer
            SetTimer this.dialogCheckTimer, 0
        ; Create new timer to check if dialog is still open
        this.dialogCheckTimer := () => this.CheckDialogStatus()
        SetTimer this.dialogCheckTimer, 500
    }

    static CheckDialogStatus() { ; Check if the dialog is still open
        if !WinExist("ahk_id " this.hActvWnd) {
            ; Dialog closed - re-enable menu and stop monitoring
            this.Enabled := true
            this.lastDialogHwnd := 0  ; Reset last dialog tracking
            SetTimer this.dialogCheckTimer, 0
            this.dialogCheckTimer := ""
            this.hActvWnd := ""
            this.pathsCache := []  ; Clear cache when dialog closes
        }
    }

    static IsFileDialog(hwnd) { ; Verify if window is actually a file dialog
        if WinGetProcessName("ahk_id " hwnd) = "cmd.exe" ; Skip command prompt windows
            return false
            
        ; Initialize control detection flags
        _SysListView321 := 0
        _ToolbarWindow321 := 0
        _DirectUIHWND1 := 0
        _Edit1 := 0
        
        ; Get control list using hwnd method
        try {
            _controlList := WinGetControlsHwnd("ahk_id " hwnd)
        } catch Error as e {
            return false
        }
        
        ; Check each control's class
        for ctrlHwnd in _controlList {
            _controlClass := WinGetClass("ahk_id " ctrlHwnd)
            
            ; Use switch for cleaner control class checking
            switch _controlClass {
                case "SysListView32":
                    _SysListView321 := 1
                case "ToolbarWindow32":
                    _ToolbarWindow321 := 1
                case "DirectUIHWND":
                    _DirectUIHWND1 := 1
                case "Edit":
                    _Edit1 := 1
            }
        }
        
        ; Return true for either type of file dialog
        return (_DirectUIHWND1 && _ToolbarWindow321 && _Edit1) 
            || (_SysListView321 && _ToolbarWindow321 && _Edit1)
    }

    static GetOpenPaths() {  ; Get all open folder paths
        paths := []
        if ProcessExist("dopus.exe") { ; Get Directory Opus paths (only if DOpus is running)
            dopusPaths := this.GetDOpusPaths()
            for path in dopusPaths
                paths.Push(path)
        }
        if ProcessExist("XY64.exe") || ProcessExist("XYplorer.exe") { ; Get XYplorer paths
            xyPaths := this.GetXYplorerPaths()
            for path in xyPaths
                paths.Push(path)
        }
        explorerPaths := this.GetExplorerPaths() ; Get Explorer paths
        for path in explorerPaths
            paths.Push(path)
        return paths
    }

    static GetDOpusPaths() { ; Get paths from Directory Opus
        tempFile := A_Temp "\dopus_paths.xml"
        paths := []
        this.activeTabsCache.Clear()  ; Clear the active tabs cache
        if FileExist(tempFile)
            FileDelete(tempFile)
        dopusPath := A_ProgramFiles this.UserDopusPath 'dopusrt.exe'
        ; Note: temp file path must NOT be quoted in the /info parameter
        withSwitch := '"' dopusPath '" /info ' tempFile ',paths'
        try {
            runWait(A_ComSpec " /c " withSwitch,,"Hide") ; <--- thanks WKen !!!
            if FileExist(tempFile) && FileGetSize(tempFile) > 0 {
                xmlContent := FileRead(tempFile)
                FileDelete(tempFile)
                paths := this.ParsePathsXML(xmlContent)
            }
        }
        return paths
    }

    static GetExplorerPaths() { ; Get paths from Windows Explorer
        paths := []
        try {
            windows := ComObject("Shell.Application").Windows
            for window in windows {
                try {
                    if window.Name = "File Explorer" {
                        path := window.Document.Folder.Self.Path
                        if RegExMatch(path, "^::\{") ; Skip special shell folders
                            continue
                        paths.Push(path)
                    }
                }
                catch {
                    ; Skip windows that cause errors
                    continue
                }
            }
        }
        catch {
            ; If Shell.Application fails, return empty
        }
        return paths
    }

    static QueryXYplorerPath(xyHwnd, message) { ; Query XYplorer and get clipboard result
        try {
            savedClip := A_Clipboard
            A_Clipboard := ""
            
            ; Send the message
            this.SendXYplorerMessage(xyHwnd, message)
            
            ; Wait for response with configurable timeout
            if ClipWait(this.XYplorerClipWaitTimeout) {
                result := A_Clipboard
                A_Clipboard := savedClip
                return Trim(result)
            }
            
            A_Clipboard := savedClip
            return ""
        }
        catch as e {
            return ""
        }
    }

    static GetXYplorerPaths() { ; Get paths from XYplorer using WM_COPYDATA messaging
        paths := []
        
        ; Find all XYplorer windows
        hwndList := WinGetList("ahk_class ThunderRT6FormDC")
        
        if hwndList.Length = 0
            return paths
        
        loop hwndList.Length {
            xyHwnd := hwndList[A_Index]
            
            ; Get the currently active path (this is the active tab)
            activePath := this.QueryXYplorerPath(xyHwnd, "::copytext get('path', a);")
            if activePath && !this.PathInList(paths, activePath) {
                paths.Push(activePath)
                ; Mark this as active (it's the currently displayed path in active pane)
                this.activeTabsCache[activePath] := "1"
            }
            
            ; Get inactive pane path (if different)
            inactivePath := this.QueryXYplorerPath(xyHwnd, "::copytext get('path', i);")
            if inactivePath && !this.PathInList(paths, inactivePath) {
                paths.Push(inactivePath)
                ; Mark if this pane is the active one (it shouldn't be, but mark anyway for consistency)
                this.activeTabsCache[inactivePath] := "0"
            }
            
            ; Get other tabs from active pane - these are not the current active tab
            loop this.XYplorerMaxTabLoops {
                tabPath := this.QueryXYplorerPath(xyHwnd, "::copytext gettoken(get('Tabs_sf', '|', 'a'), " A_Index ", '|');")
                if !tabPath  ; Stop if no result
                    break
                if !this.PathInList(paths, tabPath) {
                    paths.Push(tabPath)
                    ; These tabs exist but are not the active tab
                    this.activeTabsCache[tabPath] := "0"
                }
            }
            
            ; Get tabs from inactive pane
            loop this.XYplorerMaxTabLoops {
                tabPath := this.QueryXYplorerPath(xyHwnd, "::copytext gettoken(get('Tabs_sf', '|', 'i'), " A_Index ", '|');")
                if !tabPath  ; Stop if no result
                    break
                if !this.PathInList(paths, tabPath) {
                    paths.Push(tabPath)
                    ; Tabs in inactive pane are definitely not active
                    this.activeTabsCache[tabPath] := "0"
                }
            }
        }
        
        return paths
    }

    static PathInList(pathList, pathToCheck) { ; Helper function to check if path already exists
        for existing in pathList {
            if (existing = pathToCheck)
                return true
        }
        return false
    }

    static SendXYplorerMessage(xyHwnd, message) { ; Send WM_COPYDATA message to XYplorer
        ; Based on File_Managers_Redirection script approach
        try {
            size := StrLen(message)
            
            ; Create COPYDATA structure
            COPYDATA := Buffer(A_PtrSize * 3)
            NumPut("Ptr", 4194305, COPYDATA, 0)              ; dwData - XYplorer magic number
            NumPut("UInt", size * 2, COPYDATA, A_PtrSize)    ; cbData - size in bytes
            NumPut("Ptr", StrPtr(message), COPYDATA, A_PtrSize * 2)  ; lpData - use StrPtr directly!
            
            ; Send using SendMessageTimeout for better reliability
            DllCall("User32.dll\SendMessageTimeout", "Ptr", xyHwnd, "UInt", 74, "Ptr", 0, "Ptr", COPYDATA, "UInt", 2, "UInt", 3000, "PtrP", Result:=0, "Ptr")
        }
        catch as e {
            ; Silently fail
        }
    }

    static ParsePathsXML(xmlContent) { ; Parse XML from DOpus output
        paths := []
        loop parse xmlContent, "`n", "`r" {
            if InStr(A_LoopField, "<path") {
                if RegExMatch(A_LoopField, "<path.*?>(.*?)</path>", &match) {
                    path := match[1]
                    path := StrReplace(path, "&amp;", "&")
                    path := StrReplace(path, "&lt;", "<")
                    path := StrReplace(path, "&gt;", ">")
                    path := StrReplace(path, "&apos;", "'")
                    path := StrReplace(path, "&quot;", '"')
                    
                    ; Check if this is an active tab
                    if RegExMatch(A_LoopField, "active_tab=`"(\d)`"", &tabMatch)
                        this.activeTabsCache[path] := tabMatch[1]
                        
                    paths.Push(path)
                }
            }
        }
        return paths
    }


    static ShowPathsMenu() { ; Show paths in a menu
        if !this.pathsCache.Length {
            return
        }

        ; Ensure window is active and ready
        if !WinExist("ahk_id " this.hActvWnd) {
            return
        }
        WinActivate("ahk_id " this.hActvWnd)
        Sleep 50  ; Give window time to activate

        ; Set coordinate mode before creating menu
        CoordMode "Menu", "Window"
        
        folderMenu := Menu()
        
        for path in this.pathsCache {
            ; Check if this path is marked as active
            ; For DOpus: active_tab has a value ("1" or "2" for which pane)
            ; For XYplorer: we store "1" for active, "0" for inactive
            isActive := this.activeTabsCache.Has(path) && (this.activeTabsCache[path] != "" && this.activeTabsCache[path] != "0")
            
            folderMenu.Add(path, this.MenuHandler.Bind(this))
            if isActive {
                try {
                    folderMenu.SetIcon(path, "shell32.dll", 138)  ; Blue arrow icon
                }
            }
        }
        
        folderMenu.Add()
        ; Convert custom hotkey to human-friendly format for display in menu. 
        hkVerbose :=  (inStr(this.UserHotKey, "^")?"Ctrl+":"") 
            . (inStr(this.UserHotKey, "+")?"Shift+":"") 
            . (inStr(this.UserHotKey, "!")?"Alt+":"") 
            . (inStr(this.UserHotKey, "#")?"Win+":"") 
            . (StrUpper(SubStr(this.UserHotKey, -1)))
        folderMenu.Add("Cancel " appName " -- " hkVerbose " will re-show menu", (*) => {})
        
        ; Double check window is still active before showing menu
        if WinActive("ahk_id " this.hActvWnd) {
            ; Get window position to verify menu placement
            WinGetPos &winX, &winY, &winWidth, &winHeight, "ahk_id " this.hActvWnd
            if (winX != "" && winY != "") {  ; Only show menu if we got valid window position
                folderMenu.Show(140, 81)
            }
        }
    }

    static MenuHandler(ItemName, ItemPos, MyMenu) { ; Handle menu selection
        this.SetDialogPath(ItemName)
        this.Enabled := false ; Disable menu after selection
    }

    static SetDialogPath(path) {  ; Set path in dialog
        if !(WinExist(this.hActvWnd)) {
            return
        }
        WinActivate(this.hActvWnd)
        if SubStr(path, -1) != "\" ; Make sure path ends with a backslash
            path .= "\"
        ; Set path in dialog's edit control
        ControlSetText path, "Edit1", this.hActvWnd 
        Sleep 50
        ControlSend "{Enter}", "Edit1", this.hActvWnd
        Sleep 50
        ControlFocus "Edit1", this.hActvWnd
    }
}

TraySetIcon("shell32.dll","283") ; Icon of a little rectangle like a menu.
; If icon is changed, change below in FileCreateShortcut() too.
; Tip: Right-click SysTray icon to choose "Start with Windows."

appName := StrReplace(A_ScriptName, ".ahk") ; Assign the name of this file as "appName".
qpMenu := A_TrayMenu ; Tray Menu.
qpMenu.Delete ; Remove standard, so that app name will be at the top. 
qpMenu.Add(appName, (*) => False) ; Shows name of app at top of menu.
qpMenu.Add() ; Separator.
qpMenu.AddStandard  ; Put the standard menu items back. 
qpMenu.Add() ; Separator.
qpMenu.Add("Start with Windows", (*) => StartUpQP()) ; Add menu item at the bottom.
if FileExist(A_Startup "\" appName ".lnk")
    qpMenu.Check("Start with Windows")
; This function is only accessed via the systray menu item.  It toggles adding/removing
; link to this script in Windows Start up folder.  Applies custom icon too.
qpMenu.Default := appName
StartUpQP(*) {	
    if FileExist(A_Startup "\" appName ".lnk") {
        FileDelete(A_Startup "\" appName ".lnk")
		MsgBox("" appName " will NO LONGER auto start with Windows.",, 4096)
	} Else {
        FileCreateShortcut(A_WorkingDir "\" appName ".exe", A_Startup "\" appName ".lnk"
        , A_WorkingDir, "", "", "shell32.dll", "", "283") ; Change icon if needed.
		MsgBox("" appName " will now auto start with Windows.",, 4096)
	}
    Reload()
}

QuickPath() ; Create instance