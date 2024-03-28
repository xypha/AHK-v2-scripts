; /* AHK v2 Show/Hide OS files - CONTENTS */
; Settings
; Auto-execute
;  = Show/Hide OS files
;  = End auto-execute
; User-defined functions
;  = MyNotification
;    + MyNotificationGui
;    + EndMyNotif
;  = Toggle protected operating system (OS) files
;    + ToggleOS
;    + CheckRegWrite
;    + ToggleOSCheck
;    + WindowsRefreshOrRun
;  = Launch function in new thread with SetTimer
;    + NewThread
; ChangeLog

;------------------------------------------------------------------------------
; Settings

#Requires AutoHotkey v2.0
#SingleInstance force
#WinActivateForce
KeyHistory 500
Persistent                 ; added for standalone AHKs to prevent auto-exit

;------------------------------------------------------------------------------
; Auto-execute
; This section should always be at the top of your script

AHKname := "AHK v2 Show/Hide OS files v1.02"

; Show notification with parameters - text; duration in milliseconds; position on screen: xAxis, yAxis; timeout by - timer (1) or sleep (0)
MyNotificationGui("Loading " AHKname, -10000, 1550, 985, 1) ; 10000ms = 10 seconds (negative number so that timer will run only once), position bottom right corner (x-axis 1550 y-axis 985) on 1920×1080 display resolution; use timer

;  = Show/Hide OS files

A_TrayMenu.Delete                             ; Delete standard menu
A_TrayMenu.Add "&Toggle OS files", ToggleOS   ; User-defined function
A_TrayMenu.Add                                ; Add a separator
A_TrayMenu.AddStandard                        ; Restore standard menu
ToggleOSCheck()                               ; Query registry and check/uncheck

;  = End auto-execute

SetTimer () => EndMyNotif(), -1000 ; Reset notification timer to 1s after code in auto-execute section has finished running
Return ; Ends auto-execute

;------------------------------------------------------------------------------
; User-defined functions

;  = MyNotification

;    + MyNotificationGui

MyNotificationGui(mytext, myduration := -500, xAxis := 1550, yAxis := 985, timer := 1) {
Global MyNotification := Gui("+AlwaysOnTop -Caption +ToolWindow")   ; +ToolWindow avoids a taskbar button and an Alt-Tab menu item.
MyNotification.BackColor := "2C2C2E"                ; "2C2C2E" for dark mode ; "EEEEEE" for White background ; can be any RGB colour (it will be made transparent below)
MyNotification.SetFont("s9 w1000", "Arial")         ; font size 9, bold
MyNotification.AddText("cWhite w230 Left", mytext)  ; "cWhite" for dark mode ; use "cBlack" for black text on white background
MyNotification.Show("x1650 y985 NoActivate")        ; NoActivate avoids deactivating the currently active window
WinMove xAxis, yAxis,,, MyNotification
If timer = 1
    SetTimer () => EndMyNotif(), myduration
If timer = 0 {
    Sleep myduration * -1
    EndMyNotif()
    }
}

;--------
;    + EndMyNotif

EndMyNotif() {
MyNotification.Destroy
}

;--------
;  = Toggle protected operating system (OS) files
; inspiration from https://www.autohotkey.com/board/topic/82603-toggle-hidden-files-system-files-and-file-extensions/?p=670182

;    + ToggleOS

ToggleOS(*) {
; alternative - Run ToggleSystemFiles.bat as administrator to toggle settings - https://superuser.com/a/1151851/391770
If Status = 0 { ; enable if disabled
    RegWrite "1", "REG_DWORD", "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSuperHidden"
    CheckRegWrite(Status)
    ToggleOSCheck
    NewThread(WindowsRefreshOrRun)
    }
Else { ; disable if enabled
    RegWrite "0", "REG_DWORD", "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSuperHidden"
    CheckRegWrite(Status)
    ToggleOSCheck
    NewThread(WindowsRefreshOrRun)
    }
}

;--------
;    + CheckRegWrite

CheckRegWrite(value) { ; check If RegWrite was success
Global Status := RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSuperHidden")
If value == Status {
    MsgBox "ToggleOS Failed",, "262144" ; 262144 = Always-on-top
    ; ToolTipFn("ToggleOS Failed", -1000) ; 1s, use tooltip and exit as an alternative to MsgBox
    Exit
    }
}

;--------
;    + ToggleOSCheck

ToggleOSCheck() { ; tray tick mark
If not IsSet(Status)
    Global Status := RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSuperHidden")
If Status = 0
    A_TrayMenu.UnCheck "&Toggle OS files"
Else A_TrayMenu.Check "&Toggle OS files"
}

;--------
;    + WindowsRefreshOrRun

WindowsRefreshOrRun() {
Sleep 2000 ; 2s wait for registry change to be enforced
If WinExist("ahk_class CabinetWClass") { ; If Windows File Explorer window exists
    WinActivate
    If WinWaitActive(,, 2)      ; 2s timeout ; wait for explorer to become active window
        Send "{F5}"             ; refresh
        ; a second refresh might be needed after a few seconds to see the effects of change in settings
        ; add a Sleep command or use SetTimer prior to refresh to account for the delay
    }
Else { ; open new explorer window if one doesn't already exist ; comment out this section if not desired
    Run 'explorer.exe',,"Max"
    WinWait("ahk_class CabinetWClass",, 10) ; 10s timeout
    WinActivate
    }
}

;------------------------------------------------------------------------------
;  = Launch function in new thread with SetTimer

;    + NewThread

NewThread(Fn, timeout := -100) { ; 100ms
SetTimer () => Fn, timeout       ; use SetTimer to run commands in a new thread and prevent current thread from being paused
}

;------------------------------------------------------------------------------
; ChangeLog

/*
v1.01 - 2024.01.30
 * rename ShowSuperHidden_Status to Status for future expansion of function
 * condense ToggleOSCheck function
 * replace `HKEY_CURRENT_USER` with `HKCU` in regkeys
 * rename MyNotificationFunc to MyNotificationGui
 * remove unnecessary parentheses in If commands
 * rename Tool_TipFunc to ToolTipFn  because `Func` is a class
 * add changelog
 * add variable `AHKname` for versioning and updation of name in template and standalone scripts
 * improve comments and update headings

v1.02 - 2024.03.28
 * fix name by adding extension
 + add defaults to 'MyNotificationGui' parameters
 - remove unnecessary quotation marks "" for 'MyNotificationGui'
 * change MyNotificationGui colour scheme to white text on dark background (dark mode)
 * update MyNotificationGui duration to negative number, as convention to match ToolTipFn; consequently switch negative multiplier from SetTimer to Sleep
 * update SetTimer for more versatility
 + add 'NewThread' function
 * improve "WindowsRefreshOrRun" by adding 2s Sleep and launch using "NewThread"
 * improve 'CheckRegWrite' and 'ToggleOSCheck' - call 'RegRead' only once per 'RegWrite'
 * improve comments and update headings 
*/
