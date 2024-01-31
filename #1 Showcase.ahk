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
;    + WinClose with !RButton
;    + WinKill with ^!F4
;    + Kill All Instances Of An App with ^!+F4
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
; User-defined functions
;  = MyNotification
;    + MyNotificationGui
;    + EndMyNotif
;  = Toggle protected operating system (OS) files
;    + ToggleOS
;    + CheckRegWrite
;    + ToggleOSCheck
;    + WindowsRefreshOrRun
;  = Adjust Window Transparency
;    + GetTrans
;    + SetTransByWheel
;    + SetTransMenuFn
;    + SetTransByMenu
;  = Change Text Case
;    + ChangeCaseMenuFn
;    + ConvertLower
;    + ConvertSentence
;    + ConvertTitle
;    + ConvertUpper
;    + ConvertInvert
;    + CaseConvert
;  = Call ClipWait / Clipboard
;    + CallClipWait
;    + CallClipboard
;  = ToolTip SetTimer
;    + ToolTipFn
;  = Wrap Text In Quotes or Symbols
;    + WrapTextMenuFn
;    + WrapTextFromMenu
;    + EncText
;  = URL Encode/Decode
;    + UrlDecode
;    + UrlEncode
;  = Kill All Instances Of An App
;    + GetKillTitles
;  = Control Panel Tools
;    + ControlPanelMenuFn
;    + ControlPanelSelect
;    + List of commands
; ChangeLog

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

AHKname := "AHK v2 #1 Showcase v2.04"

; Show notification with parameters - text; duration in milliseconds; position on screen: xAxis, yAxis; timeout by - timer (1) or sleep (0)
MyNotificationGui("Loading " AHKname, "10000", "1550", "985", "1") ; 10000ms = 10 seconds, position bottom right corner (x-axis 1550 y-axis 985) on 1920×1080 display resolution; use timer

;  = Set default state of Lock keys
; Turn on/off upon startup (one-time)

SetCapsLockState "Off"   ; CapsLock   is off - Use SetCapsLockState "AlwaysOff" to force the key to stay off permanently, and uncomment `Persistent`
SetNumLockState "On"     ; NumLock    is ON
SetScrollLockState "Off" ; ScrollLock is off

;  = Show/Hide OS files

A_TrayMenu.Delete                             ; Delete standard menu
A_TrayMenu.Add "&Toggle OS files", ToggleOS   ; User-defined function
A_TrayMenu.Add                                ; Add a separator
A_TrayMenu.AddStandard                        ; Restore standard menu
ToggleOSCheck                                 ; Query registry and check/uncheck

;  = Customise Tray Icon

I_Icon := A_ScriptDir "\icons\1-512.ico"
; Icon source: https://www.iconsdb.com/caribbean-blue-icons/1-icon.html     ; CC License
; I like to number scripts 1, 2, 3... and link the scripts to Numpad shortcuts for easy editing -- see section on "Check & Reload AHK"
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
GroupAdd "FileNameSymbols"      , "Save As ahk_class #32770"                        ; Save as dialogue

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
If WinWait(A_ScriptFullPath " - AutoHotkey v" A_AhkVersion,, 3) ; wait for ListLines window to open, timeout 3s
    WinMaximize
}

^!Numpad1:: { ; Ctrl + Alt + Numpad1 keys pressed together
MyNotificationGui("Updating " AHKname, "500", "1550", "985", "0") ; use Sleep coz reload cancels timers
Reload
}

;------------------------------------------------------------------------------
;  = Remap Keys

; Disable keys that you don't use or trigger accidentally too often or become annoying
; such keys are hardware specific - desktop vs. laptop, and may vary according to region
; comment out the ones that don't work for you or don't apply to you
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

; Use Alt + Insert to toggle the 'Insert mode' and retain the key's function
!Insert::Insert     ; Source: https://gist.github.com/endolith/823381

; Note: ^Insert = Copy(^c) as Windows default - this behaviour is not changed by the above

LWin & Tab::AltTab ; Left Win key works as left Alt key - disables taskview

RAlt::!Space       ; Alt + Space brings up a window's title bar menu

^RCtrl::MButton    ; press Left & Right Ctrl button to simulate mouse Middle Click

RCtrl & Up::     Send "{PgUp}"  ; Page up   - use "&" to create 2-key combo shortcut
RCtrl & Down::   Send "{PgDn}"  ; Page down - use variable number of spaces before Send command without affecting the command itself
RControl & Left::Send "{Home}"  ; Home      - use alternate key name for RCtrl
>^Right::        Send "{End}"   ; End       - use >^ instead of Right Ctrl button and skip using "&"

!m::WinMinimize "A"         ; Alt+ M = Minimize active window

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

;------------------------------------------------------------------------------
;  = Customise CapsLock

^CapsLock::^a        ; Select all
<#CapsLock::AltTab   ; Switch windows with Right Win + CapsLock

+CapsLock:: {
SetCapsLockState "On"
MyNotificationGui("CapsLock ON", "10000", "960", "985", "1")   ; 10000ms = 10s, change to match KeyWait timeout if needed
SetTimer CapsWait, -100 ; 100ms ; add SetTimer to move KeyWait to new thread and prevent current thread from being paused
}

CapsWait() { ; runs in new thread and allows for quick toggling of CapsLock-state with +CapsLock / CapsLock / ESC keys in current thread
KeyWait "Esc", "d t10" ; hit ESC key to skip 10s timeout ; increase timeout duration to keep CapsLock ON for longer
SetCapsLockState "Off" ; Disables CapsLock immediately
MyNotification.Destroy ; and remove notification
}

CapsLock:: { ; Turn off CapsLock immediately, if on
If GetKeyState("Capslock", "T") {
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

;    + WinClose with !RButton
; Alt + right mouse button = attempt to close window

Alt & RButton:: {

MouseGetPos ,, &id
; detect the unique ID number of the window under the mouse cursor
; MouseGetPos is used instead of acitve window "A" because mouse button is used in shortcut
; The window does not have to be active to be detected. Hidden windows cannot be detected

winClass := WinGetClass("ahk_id " id) ; Retrieves the specified window's class name
If (winClass != "Shell_TrayWnd"                         ; exclude windows taskbar
 || winClass != "TopLevelWindowForOverflowXamlIsland"   ; System tray overflow window
 || winClass != "Windows.UI.Core.CoreWindow")           ; Notification Center
;|| winClass != "insert yourapp's classname")           ; uncomment to add more apps
    WinClose("ahk_id " id)  ; sends a WM_CLOSE message to the target window
    ; PostMessage 0x0112, 0xF060,,, "ahk_id " id ; alternative - same as pressing Alt+F4 or clicking a window's close button in its title bar
}

;--------
;    + WinKill with ^!F4
; Ctrl + Alt + F4 = attempt to kill window, applies to unresponsive ones when WinClose fails
; briefly tries to close the window normally and if that fails, attempts to terminate the window's process

^!F4::WinKill "A"
/* ; alternative
^!F4:: {
Process_Name := WinGetProcessName("ahk_id " id)
ProcessClose Process_Name
}
*/

;--------
;    + Kill All Instances Of An App with ^!+F4
; Ctrl + Alt + Shift + F4 = attempt to Kill All Instances Of An App

^!+F4:: {
Process_Name := WinGetProcessName("A")
Display := ( ; continuation
    "Kill all instances of this app?`n`n"    ; `n = new line
    "Name of process:`t" Process_Name        ; `t = tab
    "`n`Path of process:`t" WinGetProcessPath("ahk_exe " Process_Name)
    "`n`nNo. of visible windows: " WinGetCount("ahk_exe " Process_Name) ; no of windows ≠ no of processes
    "`nTitles of visible windows:`n    "     ; 4 spaces to add indent, but this may vary depending on system font setting
    GetKillTitles(WinGetList("ahk_exe " Process_Name)))
DetectHiddenWindows True
Display .= ( ; continuation
    "`nNo. of all windows: " WinGetCount("ahk_exe " Process_Name) " (incl. hidden)"
    "`nTitles of all windows:`n    "
    GetKillTitles(WinGetList("ahk_exe " Process_Name)))
DetectHiddenWindows False ; default
Result := MsgBox(Display, A_ScriptName " - WARNING", "Icon! YesNo Default2 262144")
; Options: add Exclamation icon; Yes or No buttons; make No the default button to prevent accidental process kill; 262144 = Always on top
If Result = "Yes"
    While ProcessExist(Process_Name)
        ProcessClose Process_Name
    ; Run A_ComSpec ' /C Taskkill /IM "' Process_Name '" /F' ; alternative - you might see a flash of command promt/terminal window. Brief explanation of flags -
    ; /C Carries out the command and then terminates - open dialogue (Win + R), paste & run "cmd.exe /?" to see other flags
    ; /IM imagename ; /F forcefully terminate - open dialogue (Win + R), paste & run "cmd.exe", paste "Taskkill /?" (without the quotation marks) and press enter to see other flags, filters and examples
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
SetTransByWheel(Trans)
}

^+WheelDown:: {         ; decreases Trans value, makes the window more transparent
Trans := GetTrans()
If Trans > 30
    Trans := Trans - 20 ; subtract 20, change for slower/faster transition
If Trans < 21
    Trans := 1          ; never set to zero, causes ERROR
SetTransByWheel(Trans)
}

F8::SetTransMenuFn

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
    If WinWaitActive( ,, 2) { ; = Sleep 2000, but sends next command as soon as activated, instead of waiting for the full 2000ms period
        Send "{F4}"
        While ControlGetClassNN(ControlGetFocus("A")) != "Microsoft.UI.Content.DesktopChildSiteBridge1" {
            Sleep 100
            If A_Index > 5 { ; = Sleep 500 ; wait until focus is on address bar, max 500ms
                ToolTipFn(A_ThisHotkey ":: Failed to focus address bar", -1000) ; 1s
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
; modified from https://www.autohotkey.com/docs/v2/lib/SendMessage.htm#ExMonitorPower

^Esc:: {
KeyWait "Esc", "T1"     ; use KeyWait instead of Sleep for faster execution
KeyWait "Control", "T1"
Sleep 100 ; wait a bit after key release in case key release would wake up the monitor again
; Sleep 1000  ; alternative to above commands
SendMessage 0x0112, 0xF170, 2,, "Program Manager"  ; 0x0112 is WM_SYSCOMMAND, 0xF170 is SC_MONITORPOWER.
}

;--------
;  = Add Control Panel Tools to a Menu

#+x::ControlPanelMenuFn   ; Win + Shift + x

;--------
;  = Change Text Case

!c::ChangeCaseMenuFn      ; Alt + C

;--------
;  = Wrap Text In Quotes or Symbols keys

#HotIf not WinActive("ahk_exe mpc-hc.exe") ; disable below hotkeys in apps that don't use it or have conflicts - Example: Media Player Classic - Home Cinema

!q::WrapTextMenuFn ; Alt + Q

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
    WinSetTitle (RegExReplace(t, Title_When_On_Top)), t
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

#z:: {
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
    Try ProcessSetPriority(LB.Text, active_pid)
    Catch ; if error
        MyNotificationGui("ERROR! Priority could not be changed!`nProcess: " Process_Name "`nPriority :  " LB.Text, "5000", "1550", "945", "1")
    Else ; if successful
        MyNotificationGui("Success! Priority changed!`nProcess: " Process_Name "`nPriority :  " LB.Text, "5000", "1550", "945", "1")
    Finally PPGui.Destroy
    }
}

;------------------------------------------------------------------------------
; #HotIf Apps
; Tailor keyboard shortcuts, commands and functions to specific windows, apps or pre-defined groups of both

;  = Firefox

#HotIf WinActive("ahk_class MozillaWindowClass") ; main window ; excludes other dialogue boxes like "Save As" originating from ahk_exe firefox.exe

; Ctrl + Shift + O = open library / bookmark manager
^+o:: {
If WinActive(" — Mozilla Firefox") ; If not new tab, then open new one
    Send "^t"
Else Send "^l"  ; If new tab, focus address bar
Sleep 500       ; wait for focus - change as per your system performance
Send "{Raw}chrome://browser/content/places/places.xhtml`n" ; `n = {Enter}
}

; Disable Ctrl + Shift + Q = Exit shortcut
^+q::Return

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
; modified from https://www.autohotkey.com/boards/viewtopic.php?p=466527&sid=6dc4a701e678a7b9ee1241ab0043ebd8#p466527

; Method #5
+WheelUp::PostMessage 0x0114, 0,, "ScrollBar1"      ; WM_HSCROLL SB_LINELEFT
+WheelDown::PostMessage 0x0114, 1,, "ScrollBar1"    ; WM_HSCROLL SB_LINERIGHT

/* ; disabled, uncomment to use - adds Loop to above commands for faster scrolling
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
If cfc1.EndReason = "Match" {
    If A_ThisHotkey = "~!" or A_ThisHotkey = "~?" ; If ! or ? is the trigger, then add a space b/w trigger and 1st character ; !a → ! A  and ?b → ? B
        Send "{BS} +" cfc1.Input
    Else {
        Send "{BS}+" cfc1.Input ; If dot or numdot is the trigger, don't add space, coz typing website address is problematic
        ; SoundBeep 1500, 50 ; play a sound when successful - Frequency(a number between 37 and 32767), Duration in milliseconds
        ; SoundPlay "C:\Windows\Media\Windows Information Bar.wav" ; alternative to SoundBeep
        }
    Exit
    }
If cfc1.EndKey = "space" { ; prevent cfc2 from firing for numbers or symbols. Example: 0.2ms is not changed to 0.2Ms
    cfc2 := InputHook("L1 V C","{Space}{LShift}{RShift}{CapsLock}", "a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,x,y,z") ; captures 2nd character, visible, case sensitive ; . a → . A
    cfc2.Start
    cfc2.Wait
    If cfc2.EndReason = "Match"
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

Esc::WinClose("A")  ; sends a WM_CLOSE message to the target window

^q::!F4             ; sends Alt + F4

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

/* ; disabled - uncomment if needed
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


/* ; disabled - uncomment if needed

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

:?*:v--:: { ; hotstring "v--"
cliptext := StrReplace(A_Clipboard,"  "," ")        ; trim double spaces
cliptext := RegExReplace(cliptext,"\s+",A_Space)    ; replace \t(tab) \s(space) \r(carriage return) \n(linefeed) with A_Space = " " (single space)
A_Clipboard := RegExReplace(cliptext," +$|^ +")     ; trim leading/trailing spaces
}

/*
Example text - Input:
Line0 ""              (blank line)
Line1 " FUBFUBFI    " (leading and trailing space)
Line2 "dvvbvvoe   df" (2+ consecutive spaces)
Line3 ""              (blank line)

Trimmed text - Output:
Line1 "FUBFUBFI dvvbvvoe df"
Explanation: blank lines are deleted, one or more of \t \s \r \n are replaced with space, resulting in Line2 being appended to Line1
*/

; Trim but keep non-blank lines

:?*:v++:: { ; hotstring "v++"
cliptext := RegExReplace(A_Clipboard,"m) +$|^ +")   ; m) = multi-line, trim leading/trailing spaces
cliptext := RegExReplace(cliptext,"`r|^`n|`n$")     ; trim \r, leading/trailing \n
A_Clipboard := RegExReplace(cliptext," +"," ")      ; trim double spaces
}

/*
Example text - Input:
Line0 ""              (blank line)
Line1 " FUBFUBFI    " (leading and trailing space)
Line2 "dvvbvvoe   df" (2+ consecutive spaces)
Line3 ""              (blank line)

Trimmed text - Output:
Line1 "FUBFUBFI"
Line2 "dvvbvvoe df"
Explanation: blank lines are deleted and spaces are trimmed, but non-blank lines are kept
*/

;------------------------------------------------------------------------------
;  = Date & Time

;    + Format Date / Time

:*x:d++::      Send FormatTime(, "yyyy.MM.dd")          ; sends 2021.02.31
:*x:date+::    Send FormatTime(, "dd.MM.yyyy")          ; sends 28.03.2020
:*x:time+::    Send FormatTime(, "h:mm tt")             ; sends 6:48 PM
:*x:datetime+::Send FormatTime(, "dd/MM/yyyy h:mm tt")  ; sends 28/03/2020 6:46 PM

;------------------------------------------------------------------------------
;  = URL Encode/Decode

:*x:url+::Send UrlEncode(A_Clipboard)

/* Encode url
    Example: https://www.google.com/
    Copy example url to clipboard
    Triger `UrlEncode` function by typing `url+`
    Output: https%3A%2F%2Fwww.google.com%2F
*/

:*x:url-::Send UrlDecode(A_Clipboard)

/* Decode url
    Example: https%3A%2F%2Fwww.google.com%2F
    Copy example url to clipboard
    Triger `UrlDecode` function by typing `url-`
    Output: https://www.google.com/
*/

;------------------------------------------------------------------------------
; User-defined functions

;  = MyNotification

;    + MyNotificationGui

MyNotificationGui(mytext, myduration, xAxis, yAxis, timer) {       ; search for `ToolTipFn` for alternative
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
ToggleOSCheck
If Status = 0 { ; enable if disabled
    RegWrite "1", "REG_DWORD", "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSuperHidden"
    CheckRegWrite(Status)
    ToggleOSCheck
    WindowsRefreshOrRun
    }
Else { ; disable if enabled
    RegWrite "0", "REG_DWORD", "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSuperHidden"
    CheckRegWrite(Status)
    ToggleOSCheck
    WindowsRefreshOrRun
    }
}

;--------
;    + CheckRegWrite

CheckRegWrite(key) { ; check if RegWrite was success
If key = RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSuperHidden")
    MsgBox "ToggleOS Failed",, "262144" ; 262144 = Always-on-top
    ; ToolTipFn("ToggleOS Failed", -1000) ; 1s, use tooltip and exit as an alternative to MsgBox
    ; Exit
}

;--------
;    + ToggleOSCheck

ToggleOSCheck() { ; tray tick mark
Global Status := RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSuperHidden")
If Status = 0
    A_TrayMenu.UnCheck "&Toggle OS files"
Else A_TrayMenu.Check "&Toggle OS files"
}

;--------
;    + WindowsRefreshOrRun

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
;  = Adjust Window Transparency

;    + GetTrans

GetTrans() {
ToolTip ; disable previous tooltip if any
MouseGetPos ,, &WinID
Trans := WinGetTransparent("ahk_id " WinID)
If not Trans
    Trans := 255
Return Trans
}

;--------
;    + SetTransByWheel

SetTransByWheel(Transparency) {
ToolTip "Transparency: " Transparency
SetTimer () => ToolTip(), -500
MouseGetPos ,, &WinID
WinSetTransparent Transparency, "ahk_id " WinID
}

;--------
;    + SetTransMenuFn

SetTransMenuFn() {
SetTransMenu := Menu()
SetTransMenu.Delete
SetTransMenu.Add("&1 255 Opaque"            ,SetTransByMenu)
SetTransMenu.Add("&2 190 Translucent"       ,SetTransByMenu) ; Semi-opaque
SetTransMenu.Add("&3 125 Semi-transparent"  ,SetTransByMenu)
SetTransMenu.Add("&4  65 Nearly Invisible"  ,SetTransByMenu)
SetTransMenu.Add("&5   1 Invisible"         ,SetTransByMenu) ; never set to zero, causes ERROR
SetTransMenu.Show
}

;--------
;    + SetTransByMenu

SetTransByMenu(item, position, SetTransMenu) {
MouseGetPos ,, &WinID
Transparency := Trim(SubStr(item, 4, 3))
WinSetTransparent Transparency, "ahk_id " WinID
}

;------------------------------------------------------------------------------
;  = Change Text Case
; inspired from https://geekdrop.com/content/super-handy-AutoHotkey-ahk-script-to-change-the-case-of-text-in-line-or-wrap-text-in-quotes

;    + ChangeCaseMenuFn

ChangeCaseMenuFn() {
ChangeCaseMenu := Menu()
ChangeCaseMenu.Delete
ChangeCaseMenu.Add("&1 lower case"      ,ConvertLower)
ChangeCaseMenu.Add("&2 Sentence case"   ,ConvertSentence)
ChangeCaseMenu.Add("&3 Title Case"      ,ConvertTitle)
ChangeCaseMenu.Add("&4 UPPER CASE"      ,ConvertUpper)
ChangeCaseMenu.Add("&5 iNVERT cASE"     ,ConvertInvert)
ChangeCaseMenu.Show
}

;--------
;    + ConvertLower

ConvertLower(*) {
CallClipboard(2)
CaseConvert(StrLower(A_Clipboard))
}

;--------
;    + ConvertSentence

ConvertSentence(*) {
CallClipboard(2)
transformed := RegExReplace(StrLower(A_Clipboard), "(((^\s*|([.!?]+\s*))[a-z])|\Wi\W)", "$U1") ; Code Credit #1
CaseConvert(transformed)
}

;--------
;    + ConvertTitle

ConvertTitle(*) {
CallClipboard(2)
CaseConvert(StrTitle(A_Clipboard))
}

;--------
;    + ConvertUpper

ConvertUpper(*) {
CallClipboard(2)
CaseConvert(StrUpper(A_Clipboard))
}

;--------
;    + ConvertInvert

ConvertInvert(*) {
CallClipboard(2)
inverted := ""
Loop Parse A_Clipboard {     ; Code Credit #2
    If StrLower(A_LoopField) == A_LoopField    ; * Code Credit #3
        inverted .= StrUpper(A_LoopField)      ; *
    Else inverted .= StrLower(A_LoopField)     ; *
    }
CaseConvert(inverted)
}
; Unicode TestString      := "abcdefghijklmnopqrstuvwxyzéâäàåçêëèïîìæôöòûùÿáíóúñABCDEFGHIJKLMNOPQRSTUVWXYZÉÂÄÀÅÇÊËÈÏÎÌÆÔÖÒÛÙŸÁÍÓÚÑ"
; Unicode iNVERT cASE     := "ABCDEFGHIJKLMNOPQRSTUVWXYZÉÂÄÀÅÇÊËÈÏÎÌÆÔÖÒÛÙŸÁÍÓÚÑabcdefghijklmnopqrstuvwxyzéâäàåçêëèïîìæôöòûùÿáíóúñ"

;--------
;    + CaseConvert

CaseConvert(caseText) {
A_Clipboard := caseText
LenTemp := StrReplace(caseText, "`r`n", "`n") ; correct count for len
Len := "+{left " Strlen(LenTemp) "}"
Send "^v" Len ; Paste new text and select it
}

; Code Credit #1 NeedleRegEx pattern modified from pattern from https://www.autohotkey.com/board/topic/24431-convert-text-uppercase-lowercase-capitalized-or-inverted/?p=158295
; Code Credit #2 idea for loop from https://www.autohotkey.com/boards/viewtopic.php?p=58417#p58417
; Code Credit #3 - 3 lines of code with a comment "; *" were adapted from a (inaccurate) answer generated from a auto-query to DuckDuckGPT by KudoAI via https://greasyfork.org/en/scripts/459849-duckduckgpt

;------------------------------------------------------------------------------
;  = Call ClipWait ± Clipboard

;    + CallClipWait

CallClipWait(secs) {
If not ClipWait(secs) {
    MyNotificationGui(A_ThisHotkey ":: Clip Failed", "2000", "1550", "985", "1") ; personal preferrence coz tooltip conflict
    ; ToolTipFn(A_ThisHotkey ":: Clip Failed", -2000) ; Alternative to MyNotification
    Exit
    }
}

;--------
;    + CallClipboard

CallClipboard(secs) {
Global clipSave := ClipboardAll() ; Global = Return clipSave
A_Clipboard := ""
Send "^c"
If not ClipWait(secs) {
    MyNotificationGui(A_ThisHotkey ":: Clip Failed", "2000", "1550", "985", "1") ; personal preferrence coz tooltip conflict
    A_Clipboard := clipSave
    Exit
    }
}

;------------------------------------------------------------------------------
;  = ToolTip SetTimer

;    + ToolTipFn

ToolTipFn(mytext, myduration) {
ToolTip ; turn off any previous tooltip
ToolTip mytext
SetTimer () => ToolTip(), myduration
}

;------------------------------------------------------------------------------
;  = Wrap Text In Quotes or Symbols
; Inspired by https://geekdrop.com/content/super-handy-autohotkey-ahk-script-to-change-the-case-of-text-in-line-or-wrap-text-in-quotes
; and https://www.autohotkey.com/board/topic/9805-easy-encloseenquote/?p=61995

;    + WrapTextMenuFn

WrapTextMenuFn() {
WrapTextMenu := Menu()
WrapTextMenu.Delete
WrapTextMenu.Add("&1   `'  Single Quotation `'"     ,WrapTextFromMenu) ; single quotation '' ; ordered in decreasing frequency of use; reorder as needed
WrapTextMenu.Add("&2   `" Double Quotation `""      ,WrapTextFromMenu) ; double quotation ""
WrapTextMenu.Add("&3   (  Round Brackets )"         ,WrapTextFromMenu) ; round breackets ()
WrapTextMenu.Add("&4   [  Square Brackets ]"        ,WrapTextFromMenu) ; square brackets []
WrapTextMenu.Add("&5   {  Flower Brackets }"        ,WrapTextFromMenu) ; flower brackets {}
WrapTextMenu.Add("&6   ``  Accent/Backtick ``"      ,WrapTextFromMenu) ; accent/backtick ``
WrapTextMenu.Add("&7  `% Percent Sign `%"           ,WrapTextFromMenu) ; percent sign %%
WrapTextMenu.Add("&8   ‘  Single Comma Quotation ’" ,WrapTextFromMenu) ; single turned comma ‘’
WrapTextMenu.Add("&9   “ Double Comma Quotation ”"  ,WrapTextFromMenu) ; double turned comma “”
WrapTextMenu.Add("&0  Remove all"                   ,WrapTextFromMenu) ; remove quotes
WrapTextMenu.Show
}

;--------
;    + WrapTextFromMenu

WrapTextFromMenu(item, position, WrapTextMenu) {
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

;--------
;    + EncText

EncText(q,p) {
CallClipboard(2) ; 2s
TextStringInitial := A_Clipboard
TextString := A_Clipboard
TextString := StrReplace(TextString, "`r`n", "`n")      ; fix carriage return + line feed for Strlen
TextString := RegExReplace(TextString,'^\s+|\s+$')      ; RegEx remove leading/trailing space
TextString := RegExReplace(TextString,'^[\[`'\(\{%`"“‘]+|^``')     ;"; remove leading  ['({%"“‘`  ; customise as your needs in WrapTextMenuFn and WrapText Keys
TextString := RegExReplace(TextString,'[\]`'\)\}%`"”’]+$|``$')     ;"; remove trailing ]')}%"”’`  ; customise as your needs in WrapTextMenuFn and WrapText Keys
TextString := q TextString p
TextString := StrReplace(TextString, "`n" p, p)
Len1 := StrLen(TextString)

; If you regularly include leading/trailing spaces within quotes, comment out above RegEx and below If statements
If RegExMatch(TextStringInitial, "^\s+") {   ; If initial string has Leading space
    TextString := " " TextString             ; add Leading space to string
    Len1++                                   ; add 1 to len
    }
If RegExMatch(TextStringInitial, "\s+$") {   ; If initial string has Trailing space
    TextString .= " "                        ; append trailing space to string
    Len1++                                   ; add 1 to len
    }

Len2 := "+{left " Len1 "}"
; Send "{Raw}" TextString ; send string with quotes
A_Clipboard := TextString ; pasting from clipboard is faster than send raw, especially for long strings
Send "^v" Len2            ; paste and select textstring ; sometimes it doesn't work properly for unknown reasons :shrug:
; A_Clipboard := TextStringInitial  ; restore original text string to clipboard if desired
}

;------------------------------------------------------------------------------
;  = URL Encode/Decode
; Modified from https://www.autohotkey.com/boards/viewtopic.php?style=7&t=116056#p517193

;    + UrlDecode

UrlDecode(Url, Enc := "UTF-8") {
Pos := 1
Loop {
    Pos := RegExMatch(Url, "i)(?:%[\da-f]{2})+", &code, Pos++)
    If Pos = 0
        Break
    var := Buffer(StrLen(code[0]) // 3, 0)
    code := SubStr(code[0], 2)
    Loop Parse code, "`%"
        NumPut("UChar", Integer("0x" A_LoopField), var, A_Index - 1)
    Url := StrReplace(Url, "`%" code, StrGet(var, Enc))
    }
Return Url
}

;--------
;    + UrlEncode

UrlEncode(str, sExcepts := "-_.", enc := "UTF-8") {
hex := "00", func := "msvcrt\swprintf"
buff := Buffer(StrPut(str, enc)), StrPut(str, buff, enc)
encoded := ""
Loop {
    If not b := NumGet(buff, A_Index - 1, "UChar")
        Break
    ch := Chr(b)
    If (b >= 0x41 && b <= 0x5A      ; A-Z
        || b >= 0x61 && b <= 0x7A   ; a-z
        || b >= 0x30 && b <= 0x39   ; 0-9
        || InStr(sExcepts, ch, True))
        encoded .= ch
    Else {
        DllCall(func, "Str", hex, "Str", "%%%02X", "UChar", b, "Cdecl")
        encoded .= hex
        }
    }
Return encoded
}

;------------------------------------------------------------------------------
;  = Kill All Instances Of An App

;    + GetKillTitles

GetKillTitles(HWNDs) {
temp := ""
Loop HWNDs.Length {
    t := RegExReplace(WinGetTitle(HWNDs[A_Index]), "\s+", A_Space)   ; remove `t `r `n
    If t = "" and A_Index < 10
        tempTitle := "  " A_Index " = #No Title`n    "   ; pad space before 1-9
    Else If t = ""
        tempTitle := A_Index " = #No Title`n    "
    Else If A_Index < 10
        tempTitle := "  " A_Index " = " t "`n    "
    Else tempTitle := A_Index " = " t "`n    "

    ; rudimentary wordwrap ; for more AHK v1 sophisticted solution - https://www.autohotkey.com/boards/viewtopic.php?t=59461
    If StrLen(tempTitle) > 61
        temp .= SubStr(tempTitle, 1, 50) "-`n            -" SubStr(tempTitle, 51, 50)
    Else temp .= tempTitle
    }
Return temp
}

;------------------------------------------------------------------------------
;  = Control Panel Tools

;    + ControlPanelMenuFn

ControlPanelMenuFn() {
ControlPanelMenu := Menu() ; starts building a pop-up menu
ControlPanelMenu.Delete    ; deletes previously built pop-up menu, if any, and then starts adding items
ControlPanelMenu.Add("&1 Control Panel"                 ,ControlPanelSelect)
ControlPanelMenu.Add("&2 Installed Apps"                ,ControlPanelSelect)
ControlPanelMenu.Add("&3 Add/Remove Programs (Legacy)"  ,ControlPanelSelect)
ControlPanelMenu.Add("&4 Defragment Interface"          ,ControlPanelSelect)
ControlPanelMenu.Add("&5 Services"                      ,ControlPanelSelect)
ControlPanelMenu.Add("&6 Sound Mixer (Legacy)"          ,ControlPanelSelect)
ControlPanelMenu.Add("&7 Registry Editor"               ,ControlPanelSelect)
ControlPanelMenu.Add("&8 Resource Monitor"              ,ControlPanelSelect)
ControlPanelMenu.Add("&9 Windows Update"                ,ControlPanelSelect)
ControlPanelMenu.Add("&0 Windows version"               ,ControlPanelSelect) ; add more triggers using alphabets (abc…) or symbols (/*-+…)
ControlPanelMenu.Show
}

;--------
;    + ControlPanelSelect

ControlPanelSelect(item, position, ControlPanelMenu) {
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

;    + List of commands
/* 'Short' list of commands (several personal modifications over the years - NOT comprehensive, at all)
Modified from https://www.autohotkey.com/boards/viewtopic.php?p=24584#p24584

; Already in Win + X menu
ComObject("shell.application").ControlPanelItem("compmgmt.msc")    ; #x | Computer Management
ComObject("shell.application").ControlPanelItem("devmgmt.msc")     ; #x | Device Manager
ComObject("shell.application").ControlPanelItem("hdwwiz.cpl")      ; #x | Device Manager ; alternative
ComObject("shell.application").ControlPanelItem("diskmgmt.msc")    ; #x | Disk Management
ComObject("shell.application").ControlPanelItem("eventvwr.msc")    ; #x | Event Viewer

; Added to Control Panel Tools
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

;------------------------------------------------------------------------------
; ChangeLog

/*
v2.04 - 2024.01.31
 * improve remap keys section to show more variations of key names, symbols and formatting
 * replace `HKEY_CURRENT_USER` with `HKCU` in regkeys
 * rename ShowSuperHidden_Status to Status for future expansion of function
 * condense ToggleOSCheck function
 * improve comments and update headings

v2.03 - 2024.01.30
 * rename function names with `Func` in the name to `Fn` because `Func` is a class
 * improve Toggle Window On Top - change WinSetTitle command to apply to known variable `t` instead of "A"
 * other minor changes

v2.02 - 2024.01.29
 * improve ListLines WinWait command by using variables

v2.01 - 2024.01.28
 * rename MyNotificationFunc to MyNotificationGui
 * specify "Save As" WinTitle in `FileNameSymbols` group to avoid conflit with apps that use ahk_class #32770 for other uses, example - Notepad++ find-replace dialogue box
 * remove unnecessary parentheses in If commands
 * add exlusion for Notification Center and System tray overflow window for WinClose command
 * change WinKill command to use active window and add alternative
 * improve continuation section in ^!+F4 hotkey
 * rename GetTitles to GetKillTitles for more specificity, and move to user-defined functions
 * improve GetKillTitles - add padding, RegExreplace for \t \r \n, add wordwrap
 * improve ^!+F4 hotkey to use ProcessClose instead of Run A_ComSpec
 * rename SetTrans to SetTransByWheel for more specificity
 * rename Tool_TipFunc to ToolTipFunc
 * add link to Display Off shortcut
 * update SoundPlay  and SoundBeep commands to AHK v2 in capitalise first letter section
 * replace || with or in single line If commands
 * replace ! with not in If commands
 * rename SetTransFunc to SetTransByMenu for more specificity
 * move position of SetTransByMenu
 * rename WrapTextFunc to WrapTextFromMenu for more specificity
 * remove unnecessary code variable from UrlDecode
 * rename ControlPanelFunc to ControlPanelSelect for more specificity
 * utlise ch variable in UrlEncode
 * some minor changes
 * improve comments and update headings

v2.00 - 2024.01.27
 * add changelog
 * add variable `AHKname` for versioning and updation of name in template and standalone scripts
 * improve comments
*/
