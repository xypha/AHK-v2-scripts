; Visit AutoHotkey (AHK) version 2 (v2) help for information - https://www.autohotkey.com/docs/v2/
; Search for below commands/functions by using control + F to search on the help webpage - https://www.autohotkey.com/docs/v2/lib/

; comments begin with semi-colon ";" at start of line or space+; " ;" in middle of line
; comments can also be show like this - "/*" comment text "*/"
; and these two methods can be combined too :)

; /* AHK v2 #1 Showcase - CONTENTS */
; Settings
; Auto-execute
;  = Set default state of Lock keys
;  = Show/Hide OS files
;  = Customise Tray Icon
;  = Capitalise first letter exclusion Group
;  = Close With Esc/Q/W Group
;  = Horizontal Scrolling Group
;  = Symbols In File Names Group
;  = End auto-execute
; Hotkeys
;  = Check & Reload AHK
;  = Remap Keys
;  = Customise CapsLock
;  = Move Mouse Pointer by pixel
;  = Close or Kill an app window
;  = Adjust Window Transparency keys
;  = Recycle Bin shortcut
;  = Display Off shortcut
;  = Add Control Panel Tools to a Menu
;  = Change Text Case
;  = Wrap Text In Quotes or Symbols keys
;  = Exchange adjacent letters
;  = Toggle Window On Top
;  = Process Priority
; #HotIf Apps
;  = Firefox
;  = Windows File Explorer
;    + General
;    + Horizontal Scrolling
;    + Copy full path
;    + Copy file names without path
;    + Copy file names without extension and path
; #HotIf Groups
;  = Capitalise the first letter of a sentence
;  = Close With Esc/Q/W keys
;  = Horizontal Scrolling
;  = Symbols In File Names keys
; Hotstrings
;  = Find & Replace in Clipboard
;    + Find & Replace dot with space
;    + Find & Replace dot with space (RegEx)
;  = Trim Clipboard
;  = Date & Time
;    + Format Date / Time
;  = URL Encode/Decode
; User-defined Functions
;  = Notification Function
;  = Toggle protected operating system (OS) files Function
;  = Windows Refresh Or Run
;  = Adjust Window Transparency Function
;  = Change Text Case Function
;  = Call ClipWait / Clipboard Function
;  = ToolTip Function
;  = Wrap Text In Quotes or Symbols Function
;  = URL Encode/Decode Function
;  = Control Panel Tools Function

;------------------------------------------------------------------------------
; Settings
; Start of script code

#Requires AutoHotkey v2.0
#SingleInstance force
#WinActivateForce
KeyHistory 500
; Persistent                 ; uncomment for standalone AHKs to prevent auto-exit

;------------------------------------------------------------------------------
; Auto-execute
; This section should always be at the top of your script

; Show notification with parameters - text; duration in milliseconds; position on screen: xAxis, yAxis; timeout by - timer (1) or sleep (0)
MyNotificationFunc("Loading AHK v2 #1 Showcase", "10000", "1550", "985", "1") ; 10000 milliseconds = 10 seconds, position bottom right corner (x-axis 1550 y-axis 985) on 1920×1080 display resolution; use timer

;  = Set default state of Lock keys
; Turn on/off upon startup (one-time)

SetCapsLockState "Off"   ; CapsLock   is off - Use SetCapsLockState "AlwaysOff" to force the key to stay off permanently, and uncomment `Persistent`
SetNumLockState "On"     ; NumLock    is ON
SetScrollLockState "Off" ; ScrollLock is off

;  = Show/Hide OS files

A_TrayMenu.Delete                             ; Delete standard menu
A_TrayMenu.Add "&Toggle OS files", ToggleOS   ; User-defined Function
A_TrayMenu.Add                                ; Add a separator
A_TrayMenu.AddStandard                        ; Restore standard menu
ToggleOSCheck                                 ; Check registry value of ShowSuperHidden_Status

;  = Customise Tray Icon

I_Icon := A_ScriptDir "\icons\1-512.ico"
; Icon source: https://www.iconsdb.com/caribbean-blue-icons/1-icon.html     ; CC License
; I like to number scripts 1, 2, 3... and link the scripts to Numpad shortcuts for easy editing
If FileExist(I_Icon)
    TraySetIcon I_Icon

;  = Capitalise first letter exclusion Group

GroupAdd "CapitaliseFirstLetter", "ahk_class #32770"                                ; Save as dialogue

;  = Close With Esc/Q/W Group

; GroupAdd "CloseWithQW"          , "ahk_exe Taskmgr.exe"                             ; Windows Task Manager ; requires UIAccess
GroupAdd "CloseWithQW"          , "Window Spy for AHKv2 ahk_class AutoHotkeyGUI"    ; AHK window spy
GroupAdd "CloseWithQW"          , "Telegram ahk_class Qt51512QWindowIcon"           ; Telegram.exe window
GroupAdd "CloseWithQW"          , "ahk_class CalcFrame"                             ; classic calculator

;  = Horizontal Scrolling Group

GroupAdd "HorizontalScroll1"    , "ahk_class ApplicationFrameWindow"                ; Modern UWP apps like calc and screen snip
GroupAdd "HorizontalScroll1"    , "ahk_class MozillaWindowClass"                    ; Firefox
GroupAdd "HorizontalScroll1"    , "ahk_class SALFRAME"                              ; LibreOffice

;  = Symbols In File Names Group

GroupAdd "FileNameSymbols"      , "ahk_class CabinetWClass"                         ; Windows file explorer
GroupAdd "FileNameSymbols"      , "ahk_class EVERYTHING"                            ; Everything
GroupAdd "FileNameSymbols"      , "ahk_class #32770"                                ; Save as dialogue

;  = End auto-execute

SetTimer EndMyNotif, -1000 ; Reset notification timer to 1s after code in auto-execute section has finished running
Return ; Ends auto-execute

; Below code can be placed anywhere in your script

;------------------------------------------------------------------------------
; Hotkeys

; ^ is Control / Ctrl key
; ! is Alt key
; # is Windows / Win key
; + is Shift key

;  = Check & Reload AHK

!Numpad1:: { ; Ctrl + Numpad1 keys pressed together
ListLines
If WinWait(".ahk - AutoHotkey v",, 3) ; wait for ListLines window to open, timeout 3s
    WinMaximize
}

^!Numpad1:: { ; Ctrl + Alt + Numpad1 keys pressed together
MyNotificationFunc("Updating AHK v2 #1 Showcase", "500", "1550", "985", "0") ; use Sleep coz reload cancels timers
Reload
}

;------------------------------------------------------------------------------
;  = Remap Keys

; Disable keys that I don't use or trigger accidentally too often or become annoying
; many such keys are hardware specific - desktop vs. laptop, and regional differences
; comment out the ones that don't apply to you
$ScrollLock::               ; disable Scroll Lock ; $ prefix forces keyboard hook
$NumLock::                  ; disable Num Lock

+NumpadDot::                ; Numpad delete (Modifier key - Shift)
NumpadDel::

Insert::                    ; Insert mode
+Insert::                   ; Shift + Insert
#Insert::                   ; Win + Insert
+Numpad0::                  ; Numpad Insert
NumpadIns::
{ ; do nothing
}

; Use Alt + Insert to toggle the 'Insert mode'
!Insert::Insert     ; Source: https://gist.github.com/endolith/823381

; Note: ^Insert = Copy(^c) as Windows default - this behaviour is not changed by the above

LWin & Tab::AltTab ; Left Win key works as left Alt key - disables taskview

RAlt::!Space       ; Alt + space brings up window menu

^RCtrl::MButton    ; press Left & Right Ctrl button to simulate mouse Middle Click

RCtrl & Up::Send "{PgUp}"       ; Page up"
RCtrl & Down::Send "{PgDn}"     ; Page down
RCtrl & Left::Send "{Home}"     ; Home
RCtrl & Right::Send "{End}"     ; End

/* ; remap media keys to navigation keys - disabled, uncomment to use
Media_Play_Pause::PgUp
Media_Stop::PgDn
Media_Prev::Home
Media_Next::End

+Media_Play_Pause::+PgUp
+Media_Stop::+PgDn
+Media_Prev::+Home
+Media_Next::+End

^Media_Play_Pause::^PgUp
^Media_Stop::^PgDn
^Media_Prev::^Home
^Media_Next::^End

^+Media_Play_Pause::^+PgUp
^+Media_Stop::^+PgDn
^+Media_Prev::^+Home
^+Media_Next::^+End
*/

!m::WinMinimize "A"         ; Alt+ M = Minimize active window

;------------------------------------------------------------------------------
;  = Customise CapsLock

^CapsLock::^a        ; Select all
<#CapsLock::AltTab   ; Switch windows with Right Win + CapsLock

+CapsLock:: {
SetCapsLockState "On"
MyNotificationFunc("CapsLock ON", "10000", "960", "985", "1")   ; 10000ms = 10s, change to match KeyWait timeout if needed
SetTimer CapsWait, -100 ; 100ms ; add SetTimer to move KeyWait to new thread and prevent current thread from being paused
}

CapsWait() { ; runs in new thread and allows for quick toggling of CapsLock-state with +CapsLock / CapsLock / ESC keys in current thread
KeyWait "Esc", "d t10" ; hit ESC key to skip 10s timeout ; increase timeout duration to keep CapsLock ON for longer
SetCapsLockState "Off" ; Disables CapsLock immediately
MyNotification.Destroy ; and remove notification
}

CapsLock:: { ; Turn off CapsLock immediately, if on
If (GetKeyState("Capslock", "T")) {
    SetCapsLockState "Off"
    MyNotification.Destroy
    }
}

;------------------------------------------------------------------------------
;  = Move Mouse Pointer by pixel
; Modified from http://www.computoredge.com/AutoHotkey/Downloads/MousePrecise.ahk

#Numpad1::MouseMove -1,  1, 0, "R"    ; Win + Numpad1 (SC04F) move down left    ↓←
#Numpad2::MouseMove  0,  1, 0, "R"    ; Win + Numpad2 (SC050) move down         ↓
#Numpad3::MouseMove  1,  1, 0, "R"    ; Win + Numpad3 (SC051) move down right   ↓→
#Numpad4::MouseMove -1,  0, 0, "R"    ; Win + Numpad4 (SC04B) move left         ←
#Numpad5::MouseMove 960,540           ; Win + Numpad5 (SC04C) move center mouse • (1920×1080 display)
#Numpad6::MouseMove  1,  0, 0, "R"    ; Win + Numpad6 (SC04D) move right        →
#Numpad7::MouseMove -1, -1, 0, "R"    ; Win + Numpad7 (SC047) move up left      ↑←
#Numpad8::MouseMove  0, -1, 0, "R"    ; Win + Numpad8 (SC048) move up           ↑
#Numpad9::MouseMove  1, -1, 0, "R"    ; Win + Numpad9 (SC049) move up right     ↑→

^#m::MouseMove 960,540 ; Test mouse position

;------------------------------------------------------------------------------
;  = Close or Kill an app window
; Modified from https://superuser.com/a/1554366/391770
; other methods to quit app - https://www.xda-developers.com/how-force-quit-applications-windows/
; also, see section "Close With Esc/Q/W keys"

Alt & RButton:: { ; Alt + right mouse button ; attempt to close window
MouseGetPos ,, &id
winClass := WinGetClass("ahk_id " id)
If (winClass != "Shell_TrayWnd")   ; exclude windows taskbar
; If (winClass != "Shell_TrayWnd" or winClass != "insert yourapp classname") ; exclude other apps using "or"
    WinClose("ahk_id " id)  ; sends a WM_CLOSE message to the target window
    ; PostMessage 0x0112, 0xF060,,, "ahk_id " id ; alternative - same as pressing Alt+F4 or clicking a window's close button in its title bar
}

; Kill window with Ctrl + Alt + F4, usually unresponsive ones if WinClose fails
^!F4:: {
MouseGetPos ,, &id
WinKill ("ahk_id " id)
}

; Kill All Instances Of An App with Ctrl + Alt + Shift + F4
^!+F4:: {
Process_Name := WinGetProcessName("A")
Display := "Kill all instances of this app?`n`n"    ; `n = new line
         . "Name of process:`t" Process_Name        ; `t = tab
         . "`n`Path of process:`t" WinGetProcessPath("ahk_exe " Process_Name)
         . "`n`nNo. of visible windows: " WinGetCount("ahk_exe " Process_Name) ; no of windows ≠ no of processes
         . "`nTitles of visible windows:`n    "     ; 4 spaces to add indent
         .  GetTitles(WinGetList("ahk_exe " Process_Name))
DetectHiddenWindows True
Display .= "`nNo. of all windows: " WinGetCount("ahk_exe " Process_Name) " (incl. hidden)"
         .  "`nTitles of all windows:`n    "
         .  GetTitles(WinGetList("ahk_exe " Process_Name))
DetectHiddenWindows False ; default
Result := MsgBox(Display, A_ScriptName " - WARNING", "Icon! YesNo Default2 262144")
; add Exclamation icon ; Yes or No buttons ; make No the default button to ensure explicit consent for TaskKill ; 262144 Always on top
If Result = "Yes"
    Run A_ComSpec ' /C Taskkill /IM "' Process_Name '" /F'
    ; /C Carries out the command and then terminates
    ; open dialogue (Win + R), paste & run "cmd.exe /?" to see other flags
    ; /IM imagename ;  /F forcefully terminate
    ; open dialogue (Win + R), paste & run "cmd.exe", paste "Taskkill /?" (without the quotation marks) and press enter to see other flags, filters and examples
}

GetTitles(HWNDs) {
temp := ""
Loop HWNDs.Length {
    t := WinGetTitle(HWNDs[A_Index])
    If t = ""
        temp .= A_Index " = #No Title`n    "
    Else temp .= A_Index " = " t "`n    "
    }
    Return temp
}

;------------------------------------------------------------------------------
;  = Adjust Window Transparency keys
; Modified from https://www.autohotkey.com/board/topic/667-transparent-windows/?p=148102

^+WheelUp:: {           ; increases Trans value, makes the window more opaque
Trans := GetTrans()
If Trans < 255
    Trans := Trans + 20 ; add 20, change for slower/faster transition
If Trans >= 255
    Trans := "Off"
SetTrans(Trans)
}

^+WheelDown:: {         ; decreases Trans value, makes the window more transparent
Trans := GetTrans()
If Trans > 30
    Trans := Trans - 20 ; subtract 20, change for slower/faster transition
If Trans < 21
    Trans := 1          ; never set to zero, causes ERROR
SetTrans(Trans)
}

F8::SetTransMenuFunc

;------------------------------------------------------------------------------
;  = Recycle Bin shortcut

^Del:: {
If WinActive("Recycle Bin ahk_class CabinetWClass")         ; If windows file explorer is active and recycle bin is in the foreground, empty Bin
    FileRecycleEmpty
Else If WinExist("Recycle Bin ahk_class CabinetWClass")     ; If explorer is showing recycle bin but is in the background, activate it
    WinActivate
/* ; If explorer is open but not showing recycle bin, change to Bin (disabled, uncomment if desired)
Else If WinExist("ahk_class CabinetWClass") {
    WinActivate
    If WinWaitActive(,, 2) { ; = Sleep 2000, but sends next command as soon as activated, instead of waiting for the full 2000ms period
        Send "{F4}"
        While (ControlGetClassNN(ControlGetFocus("A")) != "Microsoft.UI.Content.DesktopChildSiteBridge1") { ; = Sleep 500 ; wait until focus is on address bar, max 500ms
            Sleep 100
            If (A_Index > 5) {
                Tool_TipFunc(A_ThisHotkey ":: Failed to focus address bar", -1000) ; 1s
                Exit
                }
            }
        Send "{Raw}::{645ff040-5081-101b-9f08-00aa002f954e}`n"
        }
    }
*/
Else Run "::{645ff040-5081-101b-9f08-00aa002f954e}"         ; If explorer is not open, then open Bin in explorer
}

;------------------------------------------------------------------------------
;  = Display Off shortcut
; modified from AHK docs

^Esc:: {
; Sleep 1000  ; Give user a chance to release keys (in case their release would wake up the monitor again)
KeyWait "Esc", "T1"     ; use KeyWait instead of Sleep for faster execution
KeyWait "Control", "T1"
Sleep 100
SendMessage 0x0112, 0xF170, 2,, "Program Manager"  ; 0x0112 is WM_SYSCOMMAND, 0xF170 is SC_MONITORPOWER.
}

;--------
;  = Add Control Panel Tools to a Menu

#+x::ControlPanelMenuFunc   ; Win & Shift & x

;--------
;  = Change Text Case

!c::ChangeCaseMenuFunc      ; Alt + C

;--------
;  = Wrap Text In Quotes or Symbols keys

#HotIf not WinActive("ahk_exe mpc-hc.exe") ; disable below in apps that don't use it or have conflicts - example: Media Player Classic - Home Cinema

!q::WrapTextMenuFunc ; Alt + Q

; WrapText Keys - Alt + number row
!1::EncText("`'","`'")      ; enclose in single quotation '' - ' U+0027 : APOSTROPHE
!2::EncText('`"','`"')      ; enclose in double quotation "" - " U+0022 : QUOTATION MARK
!3::EncText("(",")")        ; enclose in round breackets ()
!4::EncText("[","]")        ; enclose in square brackets []
!5::EncText("{","}")        ; enclose in flower brackets {}
!6::EncText("``","``")      ; enclose in accent/backtick ``
!7::EncText("%","%")        ; enclose in percent sign %%
!8::EncText("‘","’")        ; enclose in ‘’ - ‘ U+2018 LEFT & ’ U+2019 RIGHT SINGLE QUOTATION MARK {single turned comma & comma quotation mark}
!9::EncText("“","”")        ; enclose in “” - “ U+201C LEFT & ” U+201D RIGHT DOUBLE QUOTATION MARK {double turned comma & comma quotation mark}
!0::EncText("","")          ; remove above quotes

#HotIf

;------------------------------------------------------------------------------
;  = Exchange adjacent letters
; place cursor between 2 letters. The letters reverse positions - `ab|c` becomes `ac|b`.
; Modified from http://www.computoredge.com/AutoHotkey/Downloads/LetterSwap.ahk

$!l:: { ; Alt + L
Send "{Left}+{Right 2}"
CallClipboard(2) ; 2s
SwappedLetters := SubStr(A_Clipboard,2) SubStr(A_Clipboard,1,1)
Send SwappedLetters "{Left}"
A_Clipboard := clipSave ; restore Clipboard contents
}

;------------------------------------------------------------------------------
;  = Toggle Window On Top
; Modified from https://www.autohotkey.com/board/topic/94627-button-for-always-on-top/?p=596509

!t:: {                          ; Alt + t
Title_When_On_Top := "! "       ; change title "! " as required
t := WinGetTitle("A")
ExStyle := WinGetExStyle(t)
If (ExStyle & 0x8) {            ; 0x8 is WS_EX_TOPMOST
    WinSetAlwaysOnTop 0, t      ; Turn OFF and remove Title_When_On_Top
    WinSetTitle (RegExReplace(t, Title_When_On_Top)), "A"
    }
Else {
    WinSetAlwaysOnTop 1, t      ; Turn ON and add Title_When_On_Top
    WinSetTitle Title_When_On_Top t, t
    }
}

;------------------------------------------------------------------------------
;  = Process Priority
; Hit `Win + Z` to select and change the prioty level of a process
; The current priority level of a process can be seen in the Windows Task Manager.

#z:: { ; Win + Z
active_pid := WinGetPID("A")
Process_Name := WinGetProcessName("ahk_pid " active_pid)
PPGui := Gui("AlwaysOnTop +Resize -MaximizeBox +MinSize240x230", "! Set Priority")
PPGui.AddText(, "Press ESCAPE to cancel.")
PPGui.AddText(, "Window:`n" WinGetTitle("ahk_pid " active_pid) "`n`nProcess:`n" ProcessGetPath(active_pid))
PPGui.AddText(, "Double-click to set a new priority level.")
LB := PPGui.AddListBox("r5 Choose1", ["Normal","High","Low","BelowNormal","AboveNormal"])
; Realtime omitted because any process not designed to run at Realtime priority might reduce system stability if set to that level ; add Realtime to listbox if necessary
LB.OnEvent("DoubleClick", SetPriority)
PPGui.AddButton("default", "OK").OnEvent("Click", SetPriority)
PPGui.OnEvent("Escape", (*) => PPGui.Destroy)
PPGui.OnEvent("Close", (*) => PPGui.Destroy)
PPGui.Show

SetPriority(*) {
    PPGui.Hide
    Try
        ProcessSetPriority(LB.Text, active_pid)
    Catch ; if error
        MyNotificationFunc("ERROR! Priority could not be changed!`nProcess: " Process_Name "`nPriority :  " LB.Text, "5000", "1550", "945", "1")
    Else ; if successful
        MyNotificationFunc("Success! Priority changed!`nProcess: " Process_Name "`nPriority :  " LB.Text, "5000", "1550", "945", "1")
    Finally PPGui.Destroy
    }
}

;------------------------------------------------------------------------------
; #HotIf Apps
; Tailor keyboard shortcuts, commands and functions to specific windows, apps or pre-defined groups of both

;  = Firefox

#HotIf WinActive("ahk_class MozillaWindowClass") ; main window ; excludes other dialogue boxes like "Save As" from ahk_exe firefox.exe

^+o:: {      ; Ctrl + Shift + O to open library / bookmark manager
If WinActive(" — Mozilla Firefox") ; If not new tab, then open new one
    Send "^t"
Else Send "^l"  ; If new tab, focus address bar
Sleep 500       ; wait for focus - change as per your system performance
Send "{Raw} chrome://browser/content/places/places.xhtml`n" ; `n = {Enter}
}

^+q::Return     ; disable Exit shortcut

#HotIf

;--------
;  = Windows File Explorer

;    + General

#HotIf WinActive("ahk_class CabinetWClass")

F1::F2 ; disable opening help in MS edge

; Unselect multiple files/folders
; Source: https://superuser.com/questions/78891/is-there-a-keyboard-shortcut-to-unselect-in-windows-explorer
^+a::F5

;--------
;    + Horizontal Scrolling
; Modified from https://www.autohotkey.com/boards/viewtopic.php?p=466527&sid=6dc4a701e678a7b9ee1241ab0043ebd8#p466527

+WheelUp::PostMessage 0x0114, 0,, "ScrollBar1"      ; WM_HSCROLL SB_LINELEFT
+WheelDown::PostMessage 0x0114, 1,, "ScrollBar1"    ; WM_HSCROLL SB_LINERIGHT

/* disabled, uncomment to use - add Loop for faster scrolling
+WheelUp:: {
Loop 3
    PostMessage 0x0114, 0,, "ScrollBar1"
}

+WheelDown:: {
Loop 3
    PostMessage 0x0114, 1,, "ScrollBar1"
}

*/

;--------
;    + Copy full path
; Modified from https://www.autohotkey.com/boards/viewtopic.php?p=61084#p61084

^+c:: { ; Ctrl + Shift + C
CallClipboard(2) ; Timeout 2s
A_Clipboard := A_Clipboard ; change to plain text
}
; Example: C:\Program Files\Mozilla Firefox\firefox.exe

;--------
;    + Copy file names without path

!n:: { ; Alt + N
CallClipboard(2) ; Timeout 2s
A_Clipboard := RegExReplace(A_Clipboard, "\w:\\|.+\\") ; remove path
}

; Example: firefox.exe

;--------
;    + Copy file names without extension and path

^!n:: { ; Ctrl + Alt + N
CallClipboard(2) ; Timeout 2s
files := RegExReplace(A_Clipboard, "\w:\\|.+\\")        ; remove path
files := RegExReplace(files, "\.[\w]+(`r`n|`n)","`n")   ; remove ext, CR
A_Clipboard := RegExReplace(files, "\.[\w]+$")          ; remove last ext
}

; Example: firefox

#HotIf

;------------------------------------------------------------------------------
; #HotIf Groups

;  = Capitalise the first letter of a sentence
; modified from https://www.autohotkey.com/board/topic/132938-auto-capitalize-first-letter-of-sentence/?p=719739

#HotIf not WinActive("ahk_group CapitaliseFirstLetter") ; exclude 'Save As' dialogue box

; triggers ; add or disable one or more as needed
; ~NumpadEnter:: ; disabled by default because of too many false positives
; ~Enter::       ; disabled by default - uncomment to use
~NumpadDot::
~.::
~!::
~?:: {
cfc1 := InputHook("L1 V C","{Space}{LShift}{RShift}{CapsLock}", "a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,x,y,z") ; captures 1st character, visible, case sensitive ; .a → .A
cfc1.Start
cfc1.Wait
If (cfc1.EndReason = "Match") {
    If (A_ThisHotkey = "~!" || A_ThisHotkey = "~?") ; If ! or ? is the trigger, then add a space b/w trigger and 1st character ; !a → ! A  and ?b → ? B
        Send "{BS} +" cfc1.Input
    Else {
        Send "{BS}+" cfc1.Input ; If dot or numdot is the trigger, don't add space, coz typing website address is problematic
        ; Soundbeep, 1500, 50
        ; SoundPlay % "C:\Windows\Media\Windows Information Bar.wav"
        }
    Exit
    }
If cfc1.EndKey = "space" { ; prevent cfc2 from firing for numbers or symbols. Example: 0.2ms is not changed to 0.2Ms
    cfc2 := InputHook("L1 V C","{Space}{LShift}{RShift}{CapsLock}", "a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,x,y,z") ; captures 2nd character, visible, case sensitive ; . a → . A
    cfc2.Start
    cfc2.Wait
    If (cfc2.EndReason = "Match")
        Send "{BS}+" cfc2.Input
    }
}

#HotIf

; several other AHK v1 auto-capitalisation scripts are good, such as the one linked above
; and one from computoredge - http://www.computoredge.com/AutoHotkey/Downloads/AutoSentenceCap.ahk
; and many others that use different methods to achieve this goal. Try a few and see what works for you.

;------------------------------------------------------------------------------
;  = Close With Esc/Q/W keys

#HotIf WinActive("ahk_group CloseWithQW")

Esc::WinClose("A")                     ; sends a WM_CLOSE message to the target window
^q::!F4                                ; sends Alt + F4
; ^w::PostMessage 0x0112, 0xF060,,, "A"  ; same as Alt + F4 or clicking a window's close button in its title bar

#HotIf

;------------------------------------------------------------------------------
;  = Horizontal Scrolling
; One of these four methods should work in most situations. If not,
; search the internet for other methods and send me a msg if you find one that works for you.


; Method #1 - send window message(WM) directly to move scroll bar(SB) horizontally
; default method

+WheelUp::SendMessage 0x0114, 0, 0, ControlGetFocus("A")        ; scroll left - 0x114 is WM_HSCROLL, 0 is SB_LINELEFT
+WheelDown::SendMessage 0x0114, 1, 0, ControlGetFocus("A")      ; scroll right - 1 is SB_LINERIGHT ; same as Loop 1

/* (disabled - uncomment if needed)

; add additional 'Loop' command to any method to increase the speed of scrolling. For example, in method 1

+WheelUp:: {
Loop 3         ; increase/decrease the number of loops for faster/slower scrolling
; WARNING - Do not omit number because it causes infinite loop (which is BAD)
    SendMessage 0x0114, 0, 0, ControlGetFocus("A")
}

+WheelDown:: {
Loop 3
    SendMessage 0x0114, 1, 0, ControlGetFocus("A")
}

*/

;--------
; Method #2 - simulate horizontal mouse wheel action
; test whether method #2 works using Win + Shift + Wheel Up/Down keys (3-key combo),
; then add window title/class to group #1 in auto-execute section
; to enable simpler Shift + Wheel Up/Down (2-key combo) via #HotIf command
; source: https://www.AutoHotkey.com/boards/viewtopic.php?t=76415

#+WheelUp::WheelLeft
#+WheelDown::WheelRight

; Group "HorizontalScroll1" is defined in auto-execute section
#HotIf WinActive("ahk_group HorizontalScroll1")                ; group #1

; +WheelUp::WheelLeft ; doesn't work. Explanation: https://www.autohotkey.com/boards/viewtopic.php?p=551456#p551456
+WheelUp::Send "{WheelLeft}"
+WheelDown::Send "{WheelRight}"

#HotIf


/* (disabled - uncomment if needed)

;--------
; Method #3 - turn on scroll lock and send arrow keys to scroll horizontally

#HotIf WinActive("ahk_group HorizontalScroll2")                ; group 2 - not yet defined in auto-execute

+WheelUp::SendKey("{Left}")
+WheelDown::SendKey("{Right}")

SendKey(key) {
SetScrollLockState "On"
Send key
SetScrollLockState "Off"
}

#HotIf

;--------
; Method #4 - send arrow keys directly if other methods don't work

#HotIf WinActive("ahk_group HorizontalScroll3")                ; group 3 - not yet defined in auto-execute

+WheelUp::Send "{Left 3}"
+WheelDown::Send "{Right 3}"

#HotIf

;--------
; Method #5 - horizontal scrolling for windows file explorer
see the section under " + Horizontal Scrolling"

*/

;------------------------------------------------------------------------------
;  = Symbols In File Names keys

#HotIf WinActive("ahk_group FileNameSymbols")

; replace \/:*?"<>| with ＼⧸ ： ✲ ？＂＜＞｜
; comment out the ones you don't desire, like \ → ＼
; :?*:\::{U+FF3C}                     ; \ → ＼ | replace U+005C REVERSE SOLIDUS : backslash            → U+FF3C FULLWIDTH REVERSE SOLIDUS   ; disabled

:?*:/::{U+29F8}                     ; / → ⧸  | replace U+002F SOLIDUS : slash, forward slash, virgule → U+29F8 BIG SOLIDUS
:?*b0::+::{BS}{U+FF1A}              ; : → ：  | replace U+003A COLON                                  → U+FF1A FULLWIDTH COLON
:?*:*::{U+2732}                     ; * → ✲ | replace U+002A ASTERISK : star                         → U+2732 OPEN CENTRE ASTERISK
:?*:?::{U+FF1F}                     ; ? → ？ | replace U+003F QUESTION MARK                          → U+FF1F FULLWIDTH QUESTION MARK
:?*:"::{U+FF02}                     ; " → ＂ | replace U+0022 QUOTATION MARK : double quote          → U+FF02 FULLWIDTH QUOTATION MARK
:?*:<::{U+FF1C}                     ; < → ＜ | replace U+003C LESS-THAN SIGN                         → U+FF1C FULLWIDTH LESS-THAN SIGN
:?*:>::{U+FF1E}                     ; > → ＞ | replace U+003E GREATER-THAN SIGN                      → U+FF1E FULLWIDTH GREATER-THAN SIGN
:?*:|::{U+FF5C}                     ; | → ｜ | replace U+007C VERTICAL LINE : vertical bar, pipe     → U+FF5C FULLWIDTH VERTICAL LINE

; Template -
; :*:*::{U+}                     ; ? → ? | replace ?     → ?

#HotIf

;------------------------------------------------------------------------------
; Hotstrings

;  = Find & Replace in Clipboard

;    + Find & Replace dot with space

:*:.++:: { ; hotstring ".++"
A_Clipboard := StrReplace(A_Clipboard,"."," ") ; replace each dot with space
}

/*
Find text:          "ABC..def.GHI"
Replacement text:   "ABC  def GHI"
*/

;    + Find & Replace dot with space (RegEx)

:*:.r+:: { ; hotstring ".r+"
A_Clipboard := RegExReplace(A_Clipboard,"\.+"," ") ; replace one or more dots with single space
}

/*
Find text:          "ABC..def.GHI"
Replacement text:   "ABC def GHI"
*/

;------------------------------------------------------------------------------
;  = Trim Clipboard

; Trim and change multi-line text to single line
; Modified from https://www.autohotkey.com/board/topic/89839-pasting-plain-text-from-the-clipboard/?p=568613

:?*:v--:: { ; hotkey "Control Shift V"
cliptext := StrReplace(A_Clipboard,"  "," ")        ; trim double spaces
cliptext := RegExReplace(cliptext,"\s+"," ")        ; trim `t `r `n and multiple spaces
A_Clipboard := RegExReplace(cliptext," +$|^ +")     ; trim leading/trailing spaces
}

/*
Example text:
Line0 ""              (blank line)
Line1 " FUBFUBFI    " (leading and trailing space)
Line2 "dvvbvvoe   df" (2+ consecutive spaces)
Line3 ""              (blank line)

Trimmed text: (new line is replaced with space and Line2 is appended to Line1, blank lines are deleted)
Line1 "FUBFUBFI dvvbvvoe df"
*/

; Trim but keep non-blank lines

:?*:v++:: { ; hotstring "v++"
cliptext := RegExReplace(A_Clipboard,"m) +$|^ +")   ; m) = multi-line, trim leading/trailing spaces
cliptext := RegExReplace(cliptext,"`r|^`n|`n$")     ; trim CR, leading/trailing LF
A_Clipboard := RegExReplace(cliptext," +"," ")      ; trim double spaces
}

/*
Example text:
Line0 ""              (blank line)
Line1 " FUBFUBFI    " (leading and trailing space)
Line2 "dvvbvvoe   df" (2+ consecutive spaces)
Line3 ""              (blank line)

Trimmed text: (non-blank lines are kept, spaces are trimmed and blank lines are deleted)
Line1 "FUBFUBFI"
Line2 "dvvbvvoe df"
*/

;------------------------------------------------------------------------------
;  = Date & Time

;    + Format Date / Time

:*x:d++::      Send FormatTime(, "yyyy.MM.dd") ; sends 2021.02.31
:*x:date+::    Send FormatTime(, "dd.MM.yyyy") ; sends 28.03.2020
:*x:time+::    Send FormatTime(, "h:mm tt")    ; sends 6:48 PM
:*x:datetime+::Send FormatTime(, "dd/MM/yyyy h:mm tt") ; sends 28/03/2020 6:46 PM

;------------------------------------------------------------------------------
;  = URL Encode/Decode

:*x:url+::Send UrlEncode(A_Clipboard)

/* Encode url by
    Example: https://www.google.com/
    Copy example url to clipboard
    Triger `UrlEncode` function by typing `url+`
    Output: https%3A%2F%2Fwww.google.com%2F
*/

:*x:url-::Send UrlDecode(A_Clipboard)

/* Decode url by
    Example: https%3A%2F%2Fwww.google.com%2F
    Copy example url to clipboard
    Triger `UrlDecode` function by typing `url-`
    Output: https://www.google.com/
*/

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

;--------
;  = Toggle protected operating system (OS) files Function
; inspiration from https://www.autohotkey.com/board/topic/82603-toggle-hidden-files-system-files-and-file-extensions/?p=670182

ToggleOS(*) {
ToggleOSCheck
If (ShowSuperHidden_Status = 0) { ; enable if disabled
    RegWrite "1", "REG_DWORD", "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSuperHidden"
    CheckRegWrite(ShowSuperHidden_Status)
    ToggleOSCheck
    WindowsRefreshOrRun
    }
Else { ; disable if enabled
    RegWrite "0", "REG_DWORD", "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSuperHidden"
    CheckRegWrite(ShowSuperHidden_Status)
    ToggleOSCheck
    WindowsRefreshOrRun
    }
}

CheckRegWrite(key) { ; check if RegWrite was success
If key = RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSuperHidden")
    MsgBox "ToggleOS Failed",, "262144" ; 262144 = Always-on-top
    ; Tool_TipFunc("ToggleOS Failed", -1000) ; 1s, use tooltip and exit as an alternative to MsgBox
    ; Exit
}

ToggleOSCheck() { ; tray tick mark
Global ShowSuperHidden_Status
ShowSuperHidden_Status := RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSuperHidden")
If (ShowSuperHidden_Status = 0)
    A_TrayMenu.UnCheck "&Toggle OS files"
Else {
    ShowSuperHidden_Status := 1
    A_TrayMenu.Check "&Toggle OS files"
    }
}

;--------
;  = Windows Refresh Or Run

WindowsRefreshOrRun() {
If WinExist("ahk_class CabinetWClass") { ; If Windows File Explorer window exists
    WinActivate
    If WinWaitActive( ,, 2)     ; wait for explorer to become active window ; 2s timeout
        Send "{F5}"             ; refresh
        ; a second refresh might be needed after a few seconds to see the effects of change in settings
        ; add a Sleep command or use SetTimer prior to refresh to account for the delay
    }
Else { ; open new explorer window if one doesn't already exist ; comment out this section if not desired
    Run 'explorer.exe',,"Max"
    WinWait("ahk_class CabinetWClass",, 10) ; timeout 10s
    WinActivate
    }
}

;--------
;  = Adjust Window Transparency Function

GetTrans() {
ToolTip ; disable previous tooltip if any
MouseGetPos ,, &WinID
Trans := WinGetTransparent("ahk_id " WinID)
If(!Trans)
    Trans := 255
Return Trans
}

SetTrans(Transparency) {
ToolTip("Transparency: " Transparency)
SetTimer () => ToolTip(), -500
MouseGetPos ,, &WinID
WinSetTransparent Transparency, "ahk_id " WinID
}

SetTransFunc(item, position, SetTransMenu) {
MouseGetPos ,, &WinID
Trans := Trim(SubStr(item, 4, 3))
WinSetTransparent Trans, "ahk_id " WinID
}

SetTransMenuFunc() {
SetTransMenu := Menu()
SetTransMenu.Delete
SetTransMenu.Add("&1 255 Opaque"            ,SetTransFunc)
SetTransMenu.Add("&2 190 Translucent"       ,SetTransFunc) ; Semi-opaque
SetTransMenu.Add("&3 125 Semi-transparent"  ,SetTransFunc)
SetTransMenu.Add("&4  65 Nearly Invisible"  ,SetTransFunc)
SetTransMenu.Add("&5   1 Invisible"         ,SetTransFunc) ; never set to zero, causes ERROR
SetTransMenu.Show
}

;------------------------------------------------------------------------------
;  = Change Text Case Function
; inspired from https://geekdrop.com/content/super-handy-AutoHotkey-ahk-script-to-change-the-case-of-text-in-line-or-wrap-text-in-quotes

ChangeCaseMenuFunc() {
ChangeCaseMenu := Menu()
ChangeCaseMenu.Delete
ChangeCaseMenu.Add("&1 lower case"      ,ConvertLower)
ChangeCaseMenu.Add("&2 Sentence case"   ,ConvertSentence)
ChangeCaseMenu.Add("&3 Title Case"      ,ConvertTitle)
ChangeCaseMenu.Add("&4 UPPER CASE"      ,ConvertUpper)
ChangeCaseMenu.Add("&5 iNVERT cASE"     ,ConvertInvert)
ChangeCaseMenu.Show
}

ConvertLower(*) {
CallClipboard(2)
CaseConvert(StrLower(A_Clipboard))
}

ConvertSentence(*) {
CallClipboard(2)
transformed := RegExReplace(StrLower(A_Clipboard), "(((^\s*|([.!?]+\s*))[a-z])|\Wi\W)", "$U1") ; Code Credit #1
CaseConvert(transformed)
}

ConvertTitle(*) {
CallClipboard(2)
CaseConvert(StrTitle(A_Clipboard))
}

ConvertUpper(*) {
CallClipboard(2)
CaseConvert(StrUpper(A_Clipboard))
}

ConvertInvert(*) {
CallClipboard(2)
inverted := ""
Loop Parse A_Clipboard {     ; Code Credit #2
    If (StrLower(A_LoopField) == A_LoopField)  ; * Code Credit #3
        inverted .= StrUpper(A_LoopField)      ; *
    Else inverted .= StrLower(A_LoopField)     ; *
    }
CaseConvert(inverted)
}
; Unicode TestString      := "abcdefghijklmnopqrstuvwxyzéâäàåçêëèïîìæôöòûùÿáíóúñABCDEFGHIJKLMNOPQRSTUVWXYZÉÂÄÀÅÇÊËÈÏÎÌÆÔÖÒÛÙŸÁÍÓÚÑ"
; Unicode iNVERT cASE     := "ABCDEFGHIJKLMNOPQRSTUVWXYZÉÂÄÀÅÇÊËÈÏÎÌÆÔÖÒÛÙŸÁÍÓÚÑabcdefghijklmnopqrstuvwxyzéâäàåçêëèïîìæôöòûùÿáíóúñ"

CaseConvert(caseText) {
A_Clipboard := caseText
LenTemp := StrReplace(caseText, "`r`n", "`n") ; correct count for len
Len := "+{left " Strlen(LenTemp) "}"
Send "^v" ; Pastes new text
Send Len  ; and selects it
}

; Code Credit #1 NeedleRegEx pattern modified from pattern from https://www.autohotkey.com/board/topic/24431-convert-text-uppercase-lowercase-capitalized-or-inverted/?p=158295
; Code Credit #2 idea for loop from https://www.autohotkey.com/boards/viewtopic.php?p=58417#p58417
; Code Credit #3 - 3 lines of code with a comment "; *" were adapted from a (inaccurate) answer generated from a auto-query to DuckDuckGPT by KudoAI via https://greasyfork.org/en/scripts/459849-duckduckgpt

;------------------------------------------------------------------------------
;  = Call ClipWait / Clipboard Function

CallClipWait(secs) {
If !ClipWait(secs) {
    MyNotificationFunc(A_ThisHotkey ":: Clip Failed", "2000", "1550", "985", "1") ; personal preferrence coz tooltip conflict
    ; Tool_TipFunc(A_ThisHotkey ":: Clip Failed", -2000) ; Alternative to MyNotification
    Exit
    }
}

CallClipboard(secs) {
Global clipSave := ClipboardAll() ; Global = Return clipSave
A_Clipboard := ""
Send "^c"
If !ClipWait(secs) {
    MyNotificationFunc(A_ThisHotkey ":: Clip Failed", "2000", "1550", "985", "1") ; personal preferrence coz tooltip conflict
    A_Clipboard := clipSave
    clipSave := ""
    Exit
    }
}

;------------------------------------------------------------------------------
;  = ToolTip Function

ToolTipFunc(mytext, myduration) {
ToolTip ; turn off any previous tooltip
ToolTip mytext
SetTimer () => ToolTip(), myduration
}

;------------------------------------------------------------------------------
;  = Wrap Text In Quotes or Symbols Function
; Inspired by https://geekdrop.com/content/super-handy-autohotkey-ahk-script-to-change-the-case-of-text-in-line-or-wrap-text-in-quotes
; and https://www.autohotkey.com/board/topic/9805-easy-encloseenquote/?p=61995

WrapTextMenuFunc() {
WrapTextMenu := Menu()
WrapTextMenu.Delete
WrapTextMenu.Add("&1   `'  Single Quotation `'"     ,WrapTextFunc) ; single quotation '' ; ordered in decreasing frequency of use; reorder as needed
WrapTextMenu.Add("&2   `" Double Quotation `""      ,WrapTextFunc) ; double quotation ""
WrapTextMenu.Add("&3   (  Round Brackets )"         ,WrapTextFunc) ; round breackets ()
WrapTextMenu.Add("&4   [  Square Brackets ]"        ,WrapTextFunc) ; square brackets []
WrapTextMenu.Add("&5   {  Flower Brackets }"        ,WrapTextFunc) ; flower brackets {}
WrapTextMenu.Add("&6   ``  Accent/Backtick ``"      ,WrapTextFunc) ; accent/backtick ``
WrapTextMenu.Add("&7  `% Percent Sign `%"           ,WrapTextFunc) ; percent sign %%
WrapTextMenu.Add("&8   ‘  Single Comma Quotation ’" ,WrapTextFunc) ; single turned comma ‘’
WrapTextMenu.Add("&9   “ Double Comma Quotation ”"  ,WrapTextFunc) ; double turned comma “”
WrapTextMenu.Add("&0  Remove all"                   ,WrapTextFunc) ; remove quotes
WrapTextMenu.Show
}

WrapTextFunc(item, position, WrapTextMenu) {
If position = 1
    EncText("'","'")        ; enclose in single quotation '' - ' U+0027 : APOSTROPHE
Else If position = 2
    EncText('`"','`"')      ; enclose in double quotation "" - " U+0022 : QUOTATION MARK
Else If position = 3
    EncText("(",")")        ; enclose in round breackets ()
Else If position = 4
    EncText("[","]")        ; enclose in square brackets []
Else If position = 5
    EncText("{","}")        ; enclose in flower brackets {}
Else If position = 6
    EncText("``","``")      ; enclose in accent/backtick ``
Else If position = 7
    EncText("%","%")        ; enclose in percent sign %%
Else If position = 8
    EncText("‘","’")        ; enclose in ‘’ - ‘ U+2018 LEFT & ’ U+2019 RIGHT SINGLE QUOTATION MARK {single turned comma & comma quotation mark}
Else If position = 9
    EncText("“","”")        ; enclose in “” - “ U+201C LEFT & ” U+201D RIGHT DOUBLE QUOTATION MARK {double turned comma & comma quotation mark}
Else If position = 10
    EncText("","")          ; remove above quotes
}

EncText(q,p) {
CallClipboard(2) ; 2s
TextStringInitial := A_Clipboard
TextString := A_Clipboard
TextString := StrReplace(TextString, "`r`n", "`n")      ; fix carriage return + line feed for Strlen
TextString := RegExReplace(TextString,'^\s+|\s+$')      ; RegEx remove leading/trailing space
TextString := RegExReplace(TextString,'^[\[`'\(\{%`"“‘]+|^``')     ;"; remove leading  ['({%"“‘`  ; customise as your needs in WrapTextMenuFunc and WrapText Keys
TextString := RegExReplace(TextString,'[\]`'\)\}%`"”’]+$|``$')     ;"; remove trailing ]')}%"”’`  ; customise as your needs in WrapTextMenuFunc and WrapText Keys
TextString := q TextString p
TextString := StrReplace(TextString, "`n" p, p)
Len1 := StrLen(TextString)

; If you regularly include leading/trailing spaces within quotes, comment out above RegEx and below If statements
If (RegExMatch(TextStringInitial, "^\s+")) {   ; If initial string has Leading space
    TextString := " " TextString               ; add Leading space to string
    Len1++                                     ; add 1 to len
    }
If (RegExMatch(TextStringInitial, "\s+$")) {   ; If initial string has Trailing space
    TextString .= " "                          ; append trailing space to string
    Len1++                                     ; add 1 to len
    }

Len2 := "+{left " Len1 "}"
; Send "{Raw}" TextString    ; send string with quotes
A_Clipboard := TextString    ; pasting from clipboard is faster than send raw, especially for long strings
Send "^v"
Send Len2          ; and select textstring ; sometimes it doesn't work properly for unknown reasons :shrug:
; A_Clipboard := TextStringInitial  ; restore original text string to clipboard if desired
}

;------------------------------------------------------------------------------
;  = URL Encode/Decode Function
; Modified from https://www.autohotkey.com/boards/viewtopic.php?style=7&t=116056#p517193

UrlDecode(Url, Enc := "UTF-8") {
Pos := 1
Loop {
    Pos := RegExMatch(Url, "i)(?:%[\da-f]{2})+", &code, Pos++)
    If (Pos = 0)
        Break
    code := code[0]
    var := Buffer(StrLen(code) // 3, 0)
    code := SubStr(code, 2)
    Loop Parse code, "`%"
        NumPut("UChar", Integer("0x" A_LoopField), var, A_Index - 1)
    Url := StrReplace(Url, "`%" code, StrGet(var, Enc))
    }
Return Url
}

UrlEncode(str, sExcepts := "-_.", enc := "UTF-8") {
hex := "00", func := "msvcrt\swprintf"
buff := Buffer(StrPut(str, enc)), StrPut(str, buff, enc)
encoded := ""
Loop {
    If (!b := NumGet(buff, A_Index - 1, "UChar"))
        Break
    ch := Chr(b)
    ; "is alnum" is not used because it is locale dependent.
    If (b >= 0x41 && b <= 0x5A      ; A-Z
        || b >= 0x61 && b <= 0x7A   ; a-z
        || b >= 0x30 && b <= 0x39   ; 0-9
        || InStr(sExcepts, Chr(b), True))
        encoded .= Chr(b)
    Else {
        DllCall(func, "Str", hex, "Str", "%%%02X", "UChar", b, "Cdecl")
        encoded .= hex
        }
    }
Return encoded
}

;------------------------------------------------------------------------------
;  = Control Panel Tools Function

ControlPanelMenuFunc() {
ControlPanelMenu := Menu() ; starts building a pop-up menu
ControlPanelMenu.Delete    ; deletes previously built pop-up menu, if any, and then starts adding items
ControlPanelMenu.Add("&1 Control Panel"                 ,ControlPanelFunc)
ControlPanelMenu.Add("&2 Installed Apps"                ,ControlPanelFunc)
ControlPanelMenu.Add("&3 Add/Remove Programs (Legacy)"  ,ControlPanelFunc)
ControlPanelMenu.Add("&4 Defragment Interface"          ,ControlPanelFunc)
ControlPanelMenu.Add("&5 Services"                      ,ControlPanelFunc)
ControlPanelMenu.Add("&6 Sound Mixer (Legacy)"          ,ControlPanelFunc)
ControlPanelMenu.Add("&7 Registry Editor"               ,ControlPanelFunc)
ControlPanelMenu.Add("&8 Resource Monitor"              ,ControlPanelFunc)
ControlPanelMenu.Add("&9 Windows Update"                ,ControlPanelFunc)
ControlPanelMenu.Add("&0 Windows version"               ,ControlPanelFunc) ; add more triggers using alphabets (abc…) or symbols (/*-+…)
ControlPanelMenu.Show
}

ControlPanelFunc(item, position, ControlPanelMenu) {
If position = 1
    ComObject("shell.application").ControlPanelItem("control")      ; Control Panel
If position = 2
    Run 'explorer.exe "ms-settings:appsfeatures"'                   ; Installed Apps ; Modern Add/Remove Programs
If position = 3
    ComObject("shell.application").ControlPanelItem("appwiz.cpl")   ; Add/Remove Programs ; Legacy Control Panel
If position = 4
    ComObject("shell.application").ControlPanelItem("dfrgui")       ; Defragment Interface
If position = 5
    ComObject("shell.application").ControlPanelItem("services.msc") ; Services
If position = 6
    ComObject("shell.application").ControlPanelItem("sndvol")       ; Sound Mixer (Legacy)
If position = 7
    ComObject("shell.application").ControlPanelItem("regedit")      ; Registry Editor
If position = 8
    ComObject("shell.application").ControlPanelItem("resmon.exe")   ; Resource Monitor
If position = 9
    Run 'explorer.exe "ms-settings:windowsupdate"'                  ; Windows Update
If position = 10
    ComObject("shell.application").ControlPanelItem("winver")       ; Windows version
}

/*
'Short' list of commands (several personal modifications over the years - NOT comprehensive, at all)
Modified from https://www.autohotkey.com/boards/viewtopic.php?p=24584#p24584

; Already in Win + X menu
ComObject("shell.application").ControlPanelItem("compmgmt.msc")    ; #x   | Computer Management
ComObject("shell.application").ControlPanelItem("devmgmt.msc")     ; #x   | Device Manager
ComObject("shell.application").ControlPanelItem("hdwwiz.cpl")      ; #x   | Device Manager ; alternative
ComObject("shell.application").ControlPanelItem("diskmgmt.msc")    ; #x   | Disk Management
ComObject("shell.application").ControlPanelItem("eventvwr.msc")    ; #x   | Event Viewer

; Added to Control Panel Tools function
ComObject("shell.application").ControlPanelItem("control")         ; Control Panel
Run 'explorer.exe "ms-settings:appsfeatures"'                      ; Installed Apps ; Modern Add/Remove Programs
ComObject("shell.application").ControlPanelItem("appwiz.cpl")      ; Add/Remove Programs ; Legacy Control Panel
ComObject("shell.application").ControlPanelItem("dfrgui")          ; Defragment Interface
ComObject("shell.application").ControlPanelItem("services.msc")    ; Services
ComObject("shell.application").ControlPanelItem("sndvol")          ; Sound Mixer (Legacy)
ComObject("shell.application").ControlPanelItem("regedit")         ; Registry Editor
ComObject("shell.application").ControlPanelItem("resmon.exe")      ; Resource Monitor
Run 'explorer.exe "ms-settings:windowsupdate"'                     ; Windows Update
ComObject("shell.application").ControlPanelItem("winver")          ; Windows version

Run '"::{21EC2020-3AEA-1069-A2DD-08002B30309D}"',,"Max"            ; Control Panel (view: small icons) ; alternate
::{26EE0668-A00A-44D7-9371-BEB064C98683}                           ; Control Panel (view: category) ; alternate
::{20D04FE0-3AEA-1069-A2D8-08002B30309D}                           ; This PC
Run 'SnippingTool.exe'                                             ; Snipping Tool ; alternate ; Opens modern app
Run 'rundll32 sysdm.cpl`,EditEnvironmentVariables'                 ; Environmental Variables
ComObject("shell.application").ControlPanelItem("powercfg.cpl")    ; Power Configuration ; opens Power Options
ComObject("shell.application").ControlPanelItem("msinfo32")        ; System Information
ComObject("shell.application").ControlPanelItem("timedate.cpl")    ; Date and Time Properties
ComObject("shell.application").ControlPanelItem("ncpa.cpl")        ; Network Connections
ComObject("shell.application").ControlPanelItem("mmsys.cpl")       ; Sounds and Audio ; Opens old Sound panel - Playback, Recording, Sounds, Communications
ComObject("shell.application").ControlPanelItem("dcomcnfg")        ; Component Services
ComObject("shell.application").ControlPanelItem("gpedit.msc")      ; Group Policy Editor ; N/A in Home
ComObject("shell.application").ControlPanelItem("iexplore")        ; Internet Explorer ; Opens Edge browser
ComObject("shell.application").ControlPanelItem("inetcpl.cpl")     ; Internet Properties
ComObject("shell.application").ControlPanelItem("secpol.msc")      ; Local Security Settings ; N/A in Home
ComObject("shell.application").ControlPanelItem("lusrmgr.msc")     ; Local Users and Groups ; N/A in Win10 & later?
ComObject("shell.application").ControlPanelItem("logoff")          ; Logs You Out Of Windows
ComObject("shell.application").ControlPanelItem("main.cpl")        ; Mouse Properties
ComObject("shell.application").ControlPanelItem("perfmon.msc")     ; Performance Monitor
ComObject("shell.application").ControlPanelItem("intl.cpl")        ; Regional Settings
ComObject("shell.application").ControlPanelItem("mstsc")           ; Remote Desktop ; N/A in Home
ComObject("shell.application").ControlPanelItem("wscui.cpl")       ; Security and Maintenance
ComObject("shell.application").ControlPanelItem("fsmgmt.msc")      ; Shared Folders/MMC
ComObject("shell.application").ControlPanelItem("shutdown")        ; Shuts Down Windows
ComObject("shell.application").ControlPanelItem("StikyNot")        ; Sticky Note ; N/A
ComObject("shell.application").ControlPanelItem("msconfig")        ; System Configuration Utility
ComObject("shell.application").ControlPanelItem("sysdm.cpl")       ; System Properties
ComObject("shell.application").ControlPanelItem("taskmgr")         ; Task Manager
ComObject("shell.application").ControlPanelItem("netplwiz")        ; User Accounts
ComObject("shell.application").ControlPanelItem("utilman")         ; Modern Settings App > Accessibility
ComObject("shell.application").ControlPanelItem("firewall.cpl")    ; Windows Defender Firewall
ComObject("shell.application").ControlPanelItem("wf.msc")          ; Windows Defender Firewall with Advanced Security
ComObject("shell.application").ControlPanelItem("wmimgmt.msc")     ; Windows Management Instrumentation (WMI)
ComObject("shell.application").ControlPanelItem("wuapp")           ; Windows Update App Manager ; N/A
ComObject("shell.application").ControlPanelItem("write")           ; Wordpad ; N/A
ComObject("shell.application").ShutdownWindows()                   ; Shutdown Menu

System Configuration Tools (skipped items already in #x or listed above)

C:\WINDOWS\System32\UserAccountControlSettings.exe                 ; Change User Account Control Settings
C:\WINDOWS\System32\control.exe /name Microsoft.Troubleshooting    ; Modern Settings App > System > Troubleshoot
C:\WINDOWS\System32\control.exe system                             ; Modern Settings App > System > About
C:\WINDOWS\System32\cmd.exe /k %windir%\system32\ipconfig.exe      ; View and configure network address settings
C:\WINDOWS\System32\taskmgr.exe /7                                 ; View details about programs and processes running on your computer
C:\WINDOWS\System32\cmd.exe                                        ; Open a command prompt window
C:\WINDOWS\System32\msra.exe                                       ; Receive help from  (or offer help to) a friend over the Internet
C:\WINDOWS\System32\rstrui.exe                                     ; Restore your computer system to an earlier state
C:\WINDOWS\System32\regedt32.exe                                   ; Make changes to the Windows registry
C:\WINDOWS\System32\resmon.exe                                     ; Monitor the performance and resource usage of the local computer


Find more shortcuts to various sections within modern Settings app - https://winaero.com/ms-settings-commands-in-windows-10/
Shell:AppsFolder ; shortcuts to all apps in start menu
PostMessage 0x0111, 65305,,, "C:\YourScript.ahk ahk_class AutoHotkey" ; Suspend, Toggle
PostMessage 0x0111, 65306,,, "ScriptFileName.ahk - AutoHotkey" ; Pause, Toggle
PostMessage 0x0111, 65303,,, "ScriptFileName.ahk - AutoHotkey"  ; Reload.
*/

; End of script code
