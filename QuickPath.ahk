﻿#Requires AutoHotkey v2.0
#SingleInstance Force

;===============================================================================
; Title:    QuickPath.  
; Version:  12-18-2024
; Made by:  kunkel321. 
; QuickPath has a minimal interface.  When a Windows 'Open' or 'Save as' dialog 
; is opened, whatever (in any) folders are open in DirectoryOpus, or Windows 10
; Explorer, will by listed in a popup menu.  Click a menu item to quickly add 
; that folder to the Windows dialog. Active DOpus tabs are marked with an icon.
;===============================================================================
; QuickPath was inspired by the AHK v1 app, QuickSwitch by NoNull, which has the 
; same functionality. https://www.voidtools.com/forum/viewtopic.php?f=2&t=9881
; Some of the code came from the AHK v2 app, QuickerAccess by william_ahk.
; https://www.autohotkey.com/boards/viewtopic.php?f=83&t=134379&sid
; Claude AI was used extensively, though hours of human input was needed.
;===============================================================================

TraySetIcon("shell32.dll","283") ; Icon of a little rectangle like a menu.
; If icon is changed, change below in FileCreateShortcut() too.
; Tip: Right-click SysTray icon to choose "Start with Windows."

appName := "QuickPath" ; Name it something else, if you want to.
qpMenu := A_TrayMenu ; Tray Menu.
qpMenu.Add("Start with Windows", (*) => StartUpQP()) ; Add menu item at the bottom.
if FileExist(A_Startup "\" appName ".lnk")
    qpMenu.Check("Start with Windows")
; This function is only accessed via the systray menu item.  It toggles adding/removing
; link to this script in Windows Start up folder.  Applies custom icon too.
StartUpQP(*)
{	if FileExist(A_Startup "\" appName ".lnk")
	{	FileDelete(A_Startup "\" appName ".lnk")
		MsgBox("" appName " will NO LONGER auto start with Windows.",, 4096)
	}
	Else 
	{	FileCreateShortcut(A_WorkingDir "\" appName ".exe", A_Startup "\" appName ".lnk", A_WorkingDir, "", "", "shell32.dll", "", "283")
		MsgBox("" appName " will now auto start with Windows.",, 4096)
	}
    Reload()
}


class QuickPath {
    static hActvWnd := ""  ; Handle to previously active window
    static Enabled := true  ; Flag to control whether the menu should appear
    static dialogCheckTimer := ""
    static lastCheckTime := 0  ; Time of last folder check
    static checkTimeout := 2000  ; Time to wait before checking again (2 seconds)
    static pathsCache := []  ; Cache for found paths
    static activeTabsCache := Map()  ; New: Store which paths are in active tabs
    static lastDialogHwnd := 0  ; Last dialog window we checked
    static hotkeyBound := false

    ; Initialization
    static __New() {
        ; Monitor for #32770 class windows (file dialogs)
        SetTimer () => this.CheckForFileDialog(), 100
    }

    ; Array of recognized file dialog titles - add new ones as needed
    static fileDialogs := [
        "Save As",
        "Open",
        "Save Attachment",
        "Save All Attachments"
    ]
    
    static IsFileDialog(hwnd) { ; Verify if window is actually a file dialog
        if WinGetProcessName("ahk_id " hwnd) = "cmd.exe" ; Skip command prompt windows
            return false
            
        ; Check if window title matches any known dialog types
        winTitle := WinGetTitle("ahk_id " hwnd)
        for dialogTitle in this.fileDialogs {
            if (winTitle = dialogTitle)
                return true
        }
        return false
    }

    ; Check for Open/Save dialogs and handle them
    static CheckForFileDialog() {
        dialogActive := false
        if WinActive("ahk_class #32770") {
            hwnd := WinExist("A")
            dialogActive := this.IsFileDialog(hwnd)
            if dialogActive {
                if !this.hotkeyBound { ; Ensure hotkey is bound when dialog is active
                    Hotkey "!q", (*) => this.HotkeyHandler(), "On"
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
            Hotkey "!q", "Off"
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

    ; ; Array of recognized file dialog titles - add new ones as needed
    ; static fileDialogs := [
    ;     "Save As",
    ;     "Open",
    ;     "Save Attachment",
    ;     "Save All Attachments"
    ; ]
    
    ; static IsFileDialog(hwnd) { ; Verify if window is actually a file dialog
    ;     if WinGetProcessName("ahk_id " hwnd) = "cmd.exe" ; Skip command prompt windows
    ;         return false
            
    ;     ; Check if window title matches any known dialog types
    ;     winTitle := WinGetTitle("ahk_id " hwnd)
    ;     for dialogTitle in this.fileDialogs {
    ;         if (winTitle = dialogTitle)
    ;             return true
    ;     }
    ;     return false
    ; }

    static GetOpenPaths() {  ; Get all open folder paths
        paths := []
        if ProcessExist("dopus.exe") { ; Get Directory Opus paths (only if DOpus is running)
            dopusPaths := this.GetDOpusPaths()
            for path in dopusPaths
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
        dopusPath := A_ProgramFiles '\GPSoftware\Directory Opus\dopusrt.exe'
        withSwitch := '"' dopusPath '" /info ' tempFile ',paths'
        try {
            shell := ComObject("WScript.Shell")
            exec := shell.Exec(A_ComSpec " /c " withSwitch " 2>&1")
            exec.StdOut.ReadAll() 
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
        }
        return paths
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
            if this.activeTabsCache.Has(path) {
                folderMenu.Add(path, this.MenuHandler.Bind(this))
                try {
                    folderMenu.SetIcon(path, "shell32.dll", 138)
                }
            } else {
                folderMenu.Add(path, this.MenuHandler.Bind(this))
            }
        }
        
        folderMenu.Add()
        folderMenu.Add("Cancel -- Alt+Q will re-show menu", (*) => {})
        
        ; Double check window is still active before showing menu
        if WinActive("ahk_id " this.hActvWnd) {
            ; Get window position to verify menu placement
            WinGetPos &winX, &winY, &winWidth, &winHeight, "ahk_id " this.hActvWnd
            if (winX != "" && winY != "") {  ; Only show menu if we got valid window position
                folderMenu.Show(140, 81)
            }
        }
    }


    ; static ShowPathsMenu() { ; Show paths in a menu
    ;     if !this.pathsCache.Length {
    ;         return
    ;     }
    ;     folderMenu := Menu()
    ;     for path in this.pathsCache {
    ;         ; Check if this path is in an active tab
    ;         if this.activeTabsCache.Has(path) {
    ;             ; Use shell32.dll icon for active tabs
    ;             folderMenu.Add(path, this.MenuHandler.Bind(this)) 
    ;             try {
    ;                 folderMenu.SetIcon(path, "shell32.dll", 138)
    ;             }
    ;         } 
    ;         else {
    ;             folderMenu.Add(path, this.MenuHandler.Bind(this))
    ;         }
    ;     }
    ;     folderMenu.Add()
    ;     folderMenu.Add("Cancel -- Alt+Q will re-show menu", (*) => {})
    ;     CoordMode "Menu", "Window"
    ;     folderMenu.Show(140, 81) ; <--------- Location of popup is defined here.  Relative to dialog window. 
    ; }

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

QuickPath() ; Create instance