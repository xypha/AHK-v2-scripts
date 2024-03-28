; /* AHK v2 Template - CONTENTS */
; Settings
; Auto-execute
;  = End auto-execute

; Your code goes here
; User-defined functions
;  = MyNotification
;    + MyNotificationGui
;    + EndMyNotif
; ChangeLog

;------------------------------------------------------------------------------
; Settings

#Requires AutoHotkey v2.0
#SingleInstance force
#WinActivateForce
KeyHistory 500
; Persistent                 ; uncomment for standalone AHKs to prevent auto-exit

;------------------------------------------------------------------------------
; Auto-execute
; This section should always be at the top of your script

AHKname := "AHK v2 Template v1.02"

; Show notification with parameters - text; duration in milliseconds; position on screen: xAxis, yAxis; timeout by - timer (1) or sleep (0)
MyNotificationGui("Loading " AHKname, -10000, 1550, 985, 1) ; 10000ms = 10 seconds (negative number so that timer will run only once), position bottom right corner (x-axis 1550 y-axis 985) on 1920×1080 display resolution; use timer

;  = End auto-execute

SetTimer () => EndMyNotif(), -1000 ; Reset notification timer to 1s after code in auto-execute section has finished running
Return ; Ends auto-execute

;------------------------------------------------------------------------------
; Your code can be placed anywhere in your script after auto-execute ends



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

;------------------------------------------------------------------------------
; ChangeLog

/*
v1.01 - 2024.01.30
 * rename MyNotificationFunc to MyNotificationGui
 + add changelog
 * add variable `AHKname` for versioning and updation of name in template and standalone scripts
 * improve comments and update headings
 
v1.02 - 2024.03.28
 + add defaults to 'MyNotificationGui' parameters
 - remove unnecessary quotation marks "" for 'MyNotificationGui'
 * change MyNotificationGui colour scheme to white text on dark background (dark mode)
 * update MyNotificationGui duration to negative number, as convention to match ToolTipFn; consequently switch negative multiplier from SetTimer to Sleep 
 * update SetTimer for more versatility
 * improve comments and update headings
 * improve changelog - use "fix" instead of "correct/update", use "+" for new additions and "-" for removals, "★" for new functions/sections instead of "*"
*/
