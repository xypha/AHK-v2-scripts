; /* AHK v2 Template - CONTENTS */
; Settings
; Auto-execute
;  = End auto-execute
; User-defined Functions
;  = Notification Function

;------------------------------------------------------------------------------
; Settings

#Requires AutoHotkey v2.0
#SingleInstance force
#WinActivateForce
KeyHistory 500
Persistent                                    ; added for standalone AHKs to prevent auto exit

;------------------------------------------------------------------------------
; Auto-execute
; This section should always be at the top of your script

; Show notification with parameters - text; duration in milliseconds; position on screen: xAxis, yAxis; timeout by - timer (1) or sleep (0)
MyNotificationFunc("Loading AHK v2 #1 Showcase", "10000", "1550", "985", "1") ; 10000 milliseconds = 10 seconds, position bottom right corner (x-axis 1550 y-axis 985) on 1920×1080 display resolution; use timer

;  = End auto-execute

SetTimer EndMyNotif, -1000 ; Reset notification timer to 1s after code in auto-execute section has finished running
Return ; Ends auto-execute

; Below code can be placed anywhere in your script

;------------------------------------------------------------------------------
; User-defined Functions

;  = Notification Function

MyNotificationFunc(mytext, myduration, xAxis, yAxis, timer) {       ; search for `ToolTipFunc` for alternative
Global MyNotification := Gui("+AlwaysOnTop -Caption +ToolWindow")   ; +ToolWindow avoids a taskbar button and an Alt-Tab menu item.
MyNotification.BackColor := "EEEEEE"                ; White background, can be any RGB color (it will be made transparent below)
MyNotification.SetFont("s9 w1000", "Arial")         ; font size 9, bold
MyNotification.AddText("cBlack w230 Left", mytext)  ; black text
MyNotification.Show("x1650 y985 NoActivate")        ; NoActivate avoids deactivating the currently active window
WinMove xAxis, yAxis,,, MyNotification
If timer = 1
    SetTimer EndMyNotif, myduration * -1
If timer = 0 {
    Sleep myduration
    EndMyNotif
    }
}

EndMyNotif() {
MyNotification.Destroy
}
