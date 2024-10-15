; https://github.com/xypha/AHK-v2-scripts/edit/main/No-1%20Showcase.ahk
; Last updated 2024.10.15

; Visit AutoHotkey (AHK) version 2 (v2) help for information - https://www.autohotkey.com/docs/v2/
; Search for below commands/functions by using control + F on the help webpage - https://www.autohotkey.com/docs/v2/lib/

; comments begin with semi-colon ";" at start of line or space+; " ;" in middle of line
; comments can also be show like this - "/*" comment text "*/"
; and these two methods can be combined too :)

; /* AHK v2 No-1 Showcase - CONTENTS */
; Settings
; Auto-execute
;  = Set default state of Lock keys
;  = Show/Hide OS files
;  = AHK Dark Mode
;  = Customise Tray Icon
;  = Capitalise first letter exclusion Group
;  = Close With Esc/Q/W Group
;  = Horizontal Scrolling Group
;  = Media Keys Group (disabled)
;  = Symbols In File Names Group
;  = NoWrapText Group
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
;  = Display Off shortcuts
;  = Add Control Panel Tools to a Menu
;  = Change Text Case
;  = Wrap Text In Quotes or Symbols keys
;  = Exchange adjacent letters
;  = Toggle Window On Top
;  = Process Priority
;  = Print Screen keys
; #HotIf Apps
;  = Firefox
;  = Windows File Explorer
;    + Explorer main window
;      * Unselect
;      * UnGroup
;      * Invert selection
;      * Show folder/file size in ToolTip
;      * Horizontal Scrolling
;      * Copy full path
;      * Copy file names without path
;      * Copy file names without extension and path
; #HotIf Groups
;  = Capitalise the first letter of a sentence
;  = Close With Esc/Q/W keys
;  = Horizontal Scrolling
;  = Media Keys Restored (disabled)
;  = Symbols In File Names keys
; Hotstrings
;  = Date & Time
;    + Format Date / Time
;  = URL Encode/Decode
;  = Find & Replace in Clipboard
;    + Find & Replace dot with space
;    + Find & Replace dot with space (RegEx)
;  = Trim Clipboard
; User-defined functions
;  = MyNotification
;    + MyNotificationGui
;    + EndMyNotif
;  = Toggle protected operating system (OS) files
;    + ToggleOS
;    + CheckRegWrite
;    + ToggleOSCheck
;    + WindowsRefreshOrRun
;    + RefreshExplorer
;  = Launch explorer or reuse to open path
;    + OpenFolder
;    + FocusExplorerAddressBar
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
;  = Clipboard functions
;    + CallClipWait
;    + CallClipboard
;    + CallClipboardVar
;  = ToolTip functions
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
;  = Print Screen
;    + SnipMenuFn
;    + SnipFromMenu
;    + PrintScreenFn
;    + ScreenshotFileOp
;  = Windows File Explorer Fn
;    + GetFolderSize
;    + FolderSizeFn
;    + ValidPath
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

AHKname := "AHK v2 No-1 Showcase v2.11"

; Show notification with parameters - text; duration in milliseconds; position on screen: xAxis, yAxis; timeout by - timer (1) or sleep (0)
MyNotificationGui("Loading " AHKname, -10000, 1550, 985, 1) ; 10000ms = 10 seconds (negative number so that timer will run only once), position bottom right corner (x-axis 1550 y-axis 985) on 1920×1080 display resolution; use timer

;--------
;  = Set default state of Lock keys
; Turn on/off upon startup (one-time)

SetCapsLockState    "Off"   ; CapsLock   is off - Use SetCapsLockState "AlwaysOff" to force the key to stay off permanently, and uncomment `Persistent`
SetNumLockState     "On"    ; NumLock    is ON
SetScrollLockState  "Off"   ; ScrollLock is off

;  = Show/Hide OS files

A_TrayMenu.Delete                             ; Delete standard menu
A_TrayMenu.Add "&Toggle OS files", ToggleOS   ; User-defined function
A_TrayMenu.Add                                ; Add a separator
A_TrayMenu.AddStandard                        ; Restore standard menu
ToggleOSCheck()                               ; Query registry and check/uncheck

;--------
;  = AHK Dark Mode
; download .ahk files from the `Lib` folder in this repo
; and save to disc at the same location as your script, inside a `Lib` folder 

#Include "%A_ScriptDir%\Lib\Dark Mode - ToolTip.ahk"    ; 2024.10.15
#Include "%A_ScriptDir%\Lib\Dark Mode - MsgBox.ahk"     ; 2024.10.15
; Dark Mode - Window Spy                                  ; 2024.10.15

;--------
;  = Customise Tray Icon

I_Icon := A_ScriptDir "\icons\1-512.ico"
; Icon source: https://www.iconsdb.com/caribbean-blue-icons/1-icon.html     ; CC License
; I like to number scripts 1, 2, 3... and link the scripts to Numpad shortcuts for easy editing -- see section on "Check & Reload AHK" below
If FileExist(I_Icon)
    TraySetIcon I_Icon

;--------
;  = Capitalise first letter exclusion Group

GroupAdd "CapitaliseFirstLetter", "ahk_class #32770"                                ; Save as dialogue

;--------
;  = Close With Esc/Q/W Group

; GroupAdd "CloseWithQW"          , "ahk_exe Taskmgr.exe"                             ; Windows Task Manager ; requires UIAccess
GroupAdd "CloseWithQW"          , "Window Spy for AHKv2 ahk_class AutoHotkeyGUI"    ; AHK window spy
GroupAdd "CloseWithQW"          , "ahk_class CalcFrame"                             ; classic calculator
GroupAdd "CloseWithQW"          , "Properties ahk_class #32770 ahk_exe mpc-hc.exe"  ; MediaInfo in mpc

;--------
;  = Horizontal Scrolling Group

GroupAdd "HorizontalScroll1"    , "ahk_class ApplicationFrameWindow"                ; Modern UWP apps like calc and screen snip
GroupAdd "HorizontalScroll1"    , "ahk_class MozillaWindowClass"                    ; Firefox
GroupAdd "HorizontalScroll1"    , "ahk_class SALFRAME"                              ; LibreOffice

/*
;--------
;  = Media Keys Group (disabled)
; uncomment to use, if media keys are remapped to navigation keys in "Remap Keys" section

GroupAdd "MediaKeys"            , "ahk_class MediaPlayerClassicW"                   ; MPC-HC
GroupAdd "MediaKeys"            , "MPC-HC D3D Fullscreen"                           ; MPC-HC Full screen
GroupAdd "MediaKeys"            , "ahk_class PotPlayer64"                           ; PotPlayer
*/

;--------
;  = Symbols In File Names Group

GroupAdd "FileNameSymbols"      , "ahk_class CabinetWClass"                         ; Windows file explorer
GroupAdd "FileNameSymbols"      , "ahk_class EVERYTHING"                            ; Everything
GroupAdd "FileNameSymbols"      , "Renaming ahk_exe qbittorrent.exe"                ; qBittorrent
GroupAdd "FileNameSymbols"      , "Save ahk_class #32770"                           ; Save As / Save File dialogue
GroupAdd "FileNameSymbols"      , "Export ahk_class #32770"
GroupAdd "FileNameSymbols"      , "Rename ahk_class #32770"

;--------
;  = NoWrapText Group

GroupAdd "NoWrapText"      , "ahk_exe mpc-hc.exe"                                   ; MPC-HC
GroupAdd "NoWrapText"      , "ahk_class MSPaintApp"                                 ; MS Paint (classic)

;--------
;  = End auto-execute

SetTimer () => EndMyNotif(), -1000 ; 1s ; new thread ; Reset notification timer to 1s after code in auto-execute section has finished running
Return ; Ends auto-execute

; Below code can be placed anywhere in your script

;------------------------------------------------------------------------------
; Hotkeys

; ^ is Ctrl / Control key
;   >^    is RCtrl / right control key
;   <^    is LCtrl / left control key
; ! is Alt key
;   >!    is RAlt / right ALT
;   <!    is LAlt / left ALT
;   <^>!  is ALTGr (LControl & RAlt)
; # is Windows / Win key
; + is Shift key

;  = Check & Reload AHK

!Numpad1:: { ; Alt + Numpad1 keys pressed together
ListLines
If WinWait(A_ScriptFullPath " - AutoHotkey v" A_AhkVersion,, 3) ; 3s timeout ; wait for ListLines window to open
    WinMaximize
}

^!Numpad1:: { ; Ctrl + Alt + Numpad1 keys pressed together
MyNotificationGui("Updating " AHKname,,, 985, 0) ; 500ms ; use Sleep coz reload cancels timers
Reload
}

;------------------------------------------------------------------------------
;  = Remap Keys

; Disable keys that you don't use or trigger accidentally too often or become annoying
; such keys are hardware specific - desktop vs. laptop, and may vary according to the region
; comment out the ones that don't work for you or don't apply to you
$ScrollLock::               ; disable Scroll Lock ; $ prefix forces keyboard hook
$NumLock::                  ; disable Num Lock

+NumpadDot::                ; Numpad delete (Modifier key - Shift)
NumpadDel::

Insert::                    ; Insert mode
+Insert::                   ; Shift + Insert
#Insert::                   ; Win + Insert
+Numpad0::                  ; Numpad Insert
NumpadIns:: {
} ; disable keys - do nothing

; Use Alt + Insert to toggle the 'Insert mode' and retain the key's function
!Insert::Insert     ; Source: https://gist.github.com/endolith/823381

; Note: ^Insert = Copy(^c) as Windows default - this behaviour is not changed by the above

LWin & Tab::AltTab ; Left Win key works as left Alt key - disables task view

RAlt::!Space       ; Alt + Space brings up a window's title bar menu

^RCtrl::MButton    ; press Left & Right Ctrl button to simulate mouse Middle Click

RCtrl & Up::     Send "{PgUp}"  ; Page up   - use "&" to create 2-key combo shortcut
RCtrl & Down::   Send "{PgDn}"  ; Page down - use a variable number of spaces before Send command without affecting the command itself
RControl & Left::Send "{Home}"  ; Home      - use alternate key name for RCtrl
>^Right::        Send "{End}"   ; End       - use >^ instead of Right Ctrl button and skip using "&"

!m::WinMinimize "A"         ; Alt+ M = Minimize active window
; PostMessage 0x0112, 0xF020,,, "A" ; alternative, 0x0112 = WM_SYSCOMMAND, 0xF020 = SC_MINIMIZE

/* ; remap media keys to navigation keys - disabled
; uncomment to use and enable "Media Keys Group" If required

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
; for a more comprehensive CapsLock script, visit - https://github.com/nascentt/CapShift

^CapsLock::^a        ; Select all
<#CapsLock::AltTab   ; Switch windows with Right Win + CapsLock

;--------

+CapsLock:: {
SetCapsLockState "On"
MyNotificationGui("CapsLock ON", -10000, 845)   ; 10000ms = 10s, change to match KeyWait timeout if needed
SetTimer () => CapsWait(), -100                 ; 100ms ; new thread
}

CapsWait() {
; runs in new thread and allows for quick toggling of CapsLock-state with +CapsLock / CapsLock / ESC keys in current thread
KeyWait "Esc", "d t10"          ; hit ESC key to skip 10s timeout ; increase timeout duration to keep CapsLock ON for longer
SetCapsLockState "Off"          ; Disables CapsLock immediately
MyNotification.Destroy()        ; and remove notification
}

;--------

CapsLock:: { ; Turn off CapsLock immediately, if on
If GetKeyState("CapsLock", "T") {
    SetCapsLockState "Off"
    MyNotification.Destroy()
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
; The window does not have to be active to be detected. Hidden windows cannot be detected
; WinID := WinExist("A")  ; alternative - but 'Active' window might not always be the intended target

winClass := WinGetClass("ahk_id " id) ; Retrieves the specified window's class name
If (winClass !== "Shell_TrayWnd"                         ; exclude windows taskbar
 || winClass !== "TopLevelWindowForOverflowXamlIsland"   ; System tray overflow window
 || winClass !== "Windows.UI.Core.CoreWindow"            ; Notification Center
;|| winClass !== "insert your app's window class"        ; uncomment to add more apps
    )
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
MouseGetPos ,, &id
ProcessClose WinGetProcessName("ahk_id " id)
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
    ; Run A_ComSpec ' /C Taskkill /IM /F "' Process_Name '"' ; alternative - you might see a flash of command prompt/terminal window. Brief explanation of flags -
    ; /C Carries out the command and then terminates
    ; /IM imagename
    ; /F forcefully terminates
    ; open Run dialogue (Win + R), paste "cmd.exe /?" and press OK, to see default flags
    ; then paste "Taskkill /?" (without the quotation marks) in cmd window and press enter to see 'Taskkill' specific flags, filters and examples
}

;------------------------------------------------------------------------------
;  = Adjust Window Transparency keys
; Modified from https://www.autohotkey.com/board/topic/667-transparent-windows/?p=148102

^+WheelUp:: {           ; increases Trans value, makes the window more opaque
MouseGetPos ,, &id
; id := WinExist("A")  ; alternative - but 'Active' window might not always be the intended target
Trans := GetTrans(id)
If Trans < 255
    Trans := Trans + 20 ; add 20, change for slower/faster transition
If Trans >= 255
    Trans := "Off"
SetTransByWheel(Trans, id)
}

^+WheelDown:: {         ; decreases Trans value, makes the window more transparent
MouseGetPos ,, &id
Trans := GetTrans(id)
If Trans > 20
    Trans := Trans - 20 ; subtract 20, change for slower/faster transition
Else If Trans <= 20
    Trans := 1          ; never set to zero, causes ERROR
SetTransByWheel(Trans, id)
}

F8::SetTransMenuFn

;------------------------------------------------------------------------------
;  = Recycle Bin shortcut

^Del:: {
If WinActive("Recycle Bin ahk_class CabinetWClass")         ; If windows file explorer is active and recycle bin is in the foreground, empty Bin
    FileRecycleEmpty
Else If WinExist("Recycle Bin ahk_class CabinetWClass")     ; If explorer is showing recycle bin but is in the background, activate it
    WinActivate

; use user defined function "OpenFolder" to
; (a) If an explorer is open but not showing recycle bin, change to Bin and
; (b) If an explorer is not open, then open Bin in explorer
Else OpenFolder("::{645ff040-5081-101b-9f08-00aa002f954e}") ; comment out if not desired

; alternative to OpenFolder, directly open recycle bin in a new explorer window with below command
; Else Run "::{645ff040-5081-101b-9f08-00aa002f954e}"         ; If explorer is not open, then open Bin in explorer
}

;------------------------------------------------------------------------------
;  = Display Off shortcuts
; modified from https://www.autohotkey.com/docs/v2/lib/SendMessage.htm#ExMonitorPower

>^NumpadEnter::                         ; with Right hand, press Right Ctrl + NumpadEnter keys
<^Esc:: {                               ; with Left hand, press Left Ctrl + Esc keys
If ThisHotkey == "^NumpadEnter" {
    KeyWait "Control"       , "T1"      ; use KeyWait instead of Sleep for faster execution
    KeyWait "NumpadEnter"   , "T1"
    }
Else {
    KeyWait "Esc"       , "T1"          ; use KeyWait instead of Sleep for faster execution
    KeyWait "Control"   , "T1"
    }
Sleep 100 ; wait a bit after key release to prevent key release from waking up the monitor again
; Sleep 1000  ; simpler alternative to KeyWait commands
SendMessage 0x0112, 0xF170, 2,, "Program Manager"  ; 0x0112 is WM_SYSCOMMAND, 0xF170 is SC_MONITORPOWER.
}

;--------
;  = Add Control Panel Tools to a Menu

#+x::ControlPanelMenuFn()   ; Win + Shift + x

;--------
;  = Change Text Case

!c::ChangeCaseMenuFn()      ; Alt + C

;--------
;  = Wrap Text In Quotes or Symbols keys

#HotIf not WinActive("ahk_group NoWrapText")
; disables below hotkeys in apps that belonging to this group because they don't use it or have conflicts

!q::WrapTextMenuFn ; Alt + Q

; WrapText Keys - Alt + number row
!1::EncText( "`'" , "`'")      ; enclose in single quotation '' - ' U+0027 : APOSTROPHE
!2::EncText( '`"' , '`"')      ; enclose in double quotation "" - " U+0022 : QUOTATION MARK
!3::EncText( "("  , ")" )      ; enclose in round brackets  ()
!4::EncText( "["  , "]" )      ; enclose in square brackets []
!5::EncText( "{"  , "}" )      ; enclose in flower brackets {}
!6::EncText( "``" , "``")      ; enclose in accent/backtick ``
!7::EncText( "%"  , "%" )      ; enclose in percent sign %%
!8::EncText( "‘"  , "’" )      ; enclose in ‘’ - ‘ U+2018 Left & ’ U+2019 RIGHT SINGLE QUOTATION MARK {single turned comma & comma quotation mark}
!9::EncText( "“"  , "”" )      ; enclose in “” - “ U+201C Left & ” U+201D RIGHT DOUBLE QUOTATION MARK {double turned comma & comma quotation mark}
!0::EncText( ""   , ""  )      ; remove above quotes

#HotIf

;------------------------------------------------------------------------------
;  = Exchange adjacent letters
; place cursor between 2 letters. The letters reverse positions - `ab|c` becomes `ac|b`.
; Modified from http://www.computoredge.com/AutoHotkey/Downloads/LetterSwap.ahk

$!l:: { ; Alt + L
Send "{Left}+{Right 2}"
clipped := CallClipboardVar(2) ; 2s, Exit
Send SubStr(clipped,2) SubStr(clipped,1,1) "{Left}"
; Test: AbC
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
    WinSetTitle StrReplace(t, Title_When_On_Top), t
    }
Else {
    WinSetAlwaysOnTop 1, t      ; Turn ON and add Title_When_On_Top
    WinSetTitle Title_When_On_Top t, t
    }
}

;------------------------------------------------------------------------------
;  = Process Priority
; Hit `Win + P` to select and change the priority level of a process
; The current priority level of a process can be seen in the Windows Task Manager.

#p:: { 
active_pid := WinGetPID("A")
Process_Name := WinGetProcessName("ahk_pid " active_pid)
PPGui := Gui("AlwaysOnTop +Resize -MaximizeBox +MinSize240x230", "! Set Priority")
PPGui.AddText(, "Press ESCAPE to cancel.")
PPGui.AddText(, "Window:`n" WinGetTitle("ahk_pid " active_pid) "`n`nProcess:`n" ProcessGetPath(active_pid))
PPGui.AddText(, "Double-click to set a new priority level.")
LB := PPGui.AddListBox("r5 Choose1", ["Normal","High","Low","BelowNormal","AboveNormal"])
; Realtime omitted because any process not designed to run at Realtime priority might reduce system stability if set to that level ; add Realtime to ListBox if necessary
LB.OnEvent("DoubleClick", SetPriority)
PPGui.AddButton("default", "OK").OnEvent("Click", SetPriority)
PPGui.OnEvent("Escape", (*) => PPGui.Destroy())
PPGui.OnEvent("Close", (*) => PPGui.Destroy())
PPGui.Show

SetPriority(*) {
    PPGui.Hide
    Try ProcessSetPriority(LB.Text, active_pid)
    Catch ; if error
        MyNotificationGui("ERROR! Priority could not be changed!`nProcess: " Process_Name "`nPriority :  " LB.Text, -5000) ; 5s
    Else ; if successful
        MyNotificationGui("Success! Priority changed!`nProcess: " Process_Name "`nPriority :  " LB.Text, -5000) ; 5s
    Finally PPGui.Destroy()
    }
}

;------------------------------------------------------------------------------
;  = Print Screen keys

; $PrintScreen::      ; keyboard hook $ ; commented out to preserve default function
#PrintScreen:: {    ; Win + PrintScreen
PrintScreenFn       ; take screenshot, save and rename
}

^PrintScreen::  ; Ctrl + Print Screen (key name = PrtSc, PrtScn or PrntScrn)
; #+r::         ; video snip shortcut, uncomment if desired
#+s:: {         ; Win + Shift + s
SnipMenuFn
}

;------------------------------------------------------------------------------
; #HotIf Apps
; Tailor keyboard shortcuts, commands and functions to specific windows, apps or pre-defined groups of both

;  = Firefox

#HotIf WinActive("ahk_class MozillaWindowClass") ; main window ; excludes other dialogue boxes like "Save As" originating from ahk_exe firefox.exe

; Ctrl + Shift + F = close Find Bar
^+f::Send "^f{Esc}"

; Ctrl + Shift + H = open Homepage
^+h::Send "^labout:home{Enter}" ; Go backwards and Go forwards button history is preserved, but blank grey background may be seen instead of new tab background image
; Send "^w^t"     ; alternative - Go backwards and Go forwards button history is lost, but wallpaper loads correctly

; Ctrl + Shift + O = open library / bookmark manager
^+o:: {
If WinActive(" — Mozilla Firefox") ; If not new tab, then open new one
    Send "^t"
Else Send "^l"  ; If new tab, focus address bar
Sleep 500       ; wait for focus - change as per your system performance
Send "{Raw}chrome://browser/content/places/places.xhtml`n" ; `n = {Enter}
}

; Disable Ctrl + Shift + Q = Exit (default Firefox shortcut)
^+q::Return

#HotIf

;--------
;  = Windows File Explorer

#HotIf WinActive("ahk_class CabinetWClass")

;    + Explorer main window

F1::F2 ; disable opening help in MS edge

;--------
;      * Unselect
; Unselect multiple files/folders
; Source: https://superuser.com/questions/78891/is-there-a-keyboard-shortcut-to-unselect-in-windows-explorer

^+a::F5 ; Ctrl + Shift + A = unselect by sending {F5} key ; same as Right Click > Refresh


;--------
;      * UnGroup
; Change the annoying `Group by Date modified` default in Downloads folder to `Group by Date (None)` 

^g:: {  ; Ctrl + G
MouseMove 541, 146
Send "{Click}"
If WinWait("PopupHost ahk_class Microsoft.UI.Content.PopupWindowSiteBridge ahk_exe explorer.exe",, 1) ; 1s
    Send "{Up}{Right}"
    Sleep 250
    Send "{Up}{Enter}"
}

;--------
;      * Invert selection
; For the Windows 11 23H2 Windows File Explorer

^i:: {                    ; Ctrl + I
MouseMove 750, 150        ; see Note 1
Send "{Click}"            ; click on "see more" option
If WinWait("PopupHost ahk_class Microsoft.UI.Content.PopupWindowSiteBridge ahk_exe explorer.exe",, 1) { ; 1s
; Sleep 500                 ; alternative to WinWait ; see Note 2
    Send "{Up 3}"         ; select "Invert selection" option
    ; Send "{Enter}"        ; see Note 3
    }
}

/*
Note 1: location of "see more" option on screen
MouseMove - move the mouse cursor to x,y coordinates on 'Screen' use Window Spy to determine coordinates for your own screen(s)
might want to change "CoordMode" if you have problems, visit help page: https://www.autohotkey.com/docs/v2/lib/CoordMode.htm

Note 2: wait 500ms for window to respond to keys
wait times (in milliseconds) don't work sometimes, try changing them to see what is sufficient for your PC spec/performance.
Sometimes it doesn't work at all on the 1st try (ex: when a new explorer instance is first opened after login/restart)
but works on subsequent attempts. *shrug* no idea why.

Note 3: uncomment this line after test of previous steps
or leave it as is if you prefer to do this step manually, especially if you notice inconsistent selection of options.
*/

/* ; Invert selection in older Win 11 and 7/8/8.1/10 systems
; applicable for ribbon UI enabled File Explorer
; Modified from https://www.autohotkey.com/boards/viewtopic.php?f=76&t=27564

^i::PostMessage 0x111, 28706,, "SHELLDLL_DefView1", "A"
*/

;--------
;      * Show folder/file size in ToolTip

^+s:: {

; check keyboard focus
ClassNN := ControlGetClassNN(ControlGetFocus("A"))

; If keyboard focus = file list
If ClassNN == "DirectUIHWND2"
    path := ValidPath()

; If keyboard focus = address bar
Else If ClassNN == "Microsoft.UI.Content.DesktopChildSiteBridge1" {
    path := ValidPath()
    Send "{F6 2}{Down}{Home}" ; Return focus to file list
    }

; If keyboard focus = navigation pane
Else If ClassNN == "SysTreeView321" {
    ToolTipFn("Focus is in Navigation Pane!", -2000) ; 2s
    Send "{F6}{Home}" ; Return focus to file list
    Exit
    }

; keyboard focus = unknown.. search or dialogue box or something else?
Else {
    ToolTipFn("Unknown focus!`nClassNN: " ClassNN, -2000) ; 2s
    Exit
    }

; calculate folder size and display
GetFolderSize(path[1], path[2])
}

;--------
;      * Horizontal Scrolling
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
;      * Copy full path
; Modified from https://www.autohotkey.com/boards/viewtopic.php?p=61084#p61084

^+c:: { ; Ctrl + Shift + C
CallClipboard(2) ; 2s, Exit
A_Clipboard := A_Clipboard ; change to plain text
}
; Output Example: C:\Program Files\Mozilla Firefox\firefox.exe

;--------
;      * Copy file names without path

!n:: { ; Alt + N
A_Clipboard := RegExReplace(CallClipboardVar(2), "\w:\\|.+\\") ; 2s, Exit ; remove path
}

; Output Example: firefox.exe

;--------
;      * Copy file names without extension and path

^!n:: { ; Ctrl + Alt + N
files := RegExReplace(CallClipboardVar(2), "\w:\\|.+\\")    ; 2s, Exit ; remove path
files := RegExReplace(files, "\.[\w]+(`r`n|`n)","`n")       ; remove ext, CR
A_Clipboard := RegExReplace(files, "\.[\w]+$")              ; remove last ext
}

; Output Example: firefox

#HotIf

;------------------------------------------------------------------------------
; #HotIf Groups

;  = Capitalise the first letter of a sentence
; modified from https://www.autohotkey.com/board/topic/132938-auto-capitalize-first-letter-of-sentence/?p=719739

#HotIf not WinActive("ahk_group CapitaliseFirstLetter") ; exclude 'Save As' dialogue box

; triggers ; add or disable one or more as needed
; ~NumpadEnter:: ; disabled by default because of too many false positives
; ~Enter::       ; disabled - uncomment to use
~NumpadDot::
~.::
~!::
~?:: {
cfc1 := InputHook("L1 V C","{Space}{LShift}{RShift}{CapsLock}", "a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,x,y,z") ; captures 1st character, visible, case sensitive ; .a → .A
cfc1.Start
cfc1.Wait
If cfc1.EndReason = "Match" {
    If ThisHotkey = "~!" or ThisHotkey = "~?" ; If ! or ? is the trigger, then add a space b/w trigger and 1st character ; !a → ! A  and ?b → ? B
        Send "{BS} +" cfc1.Input
    Else {
        Send "{BS}+" cfc1.Input ; If Dot or NumpadDot is the trigger, don't add space, coz typing the website address is problematic
        ; SoundBeep 1500, 50 ; play a sound when successful - Frequency(a number between 37 and 32767), Duration in milliseconds
        ; SoundPlay "C:\Windows\Media\Windows Information Bar.wav" ; alternative to SoundBeep
        ; SoundPlay A_WinDir "\Media\Windows Balloon.wav"          ; alternative to SoundBeep
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

Esc::WinClose "A"  ; sends a WM_CLOSE message to the target window

^q::!F4             ; sends Alt + F4

; ^w::PostMessage 0x0112, 0xF060,,, "A"  ; same as Alt + F4 or clicking a window's close button in its title bar

#HotIf

;------------------------------------------------------------------------------
;  = Horizontal Scrolling
; One of these four methods should work in most situations. If not,
; search the internet for other methods and send me a msg if you find one that works for you.


; Method #1 - send window message(WM) directly to move scroll bar(SB) horizontally
; default method

+WheelUp::SendMessage 0x0114, 0, 0, ControlGetFocus("A")        ; scroll Left - 0x114 is WM_HSCROLL, 0 is SB_LINELEFT
+WheelDown::SendMessage 0x0114, 1, 0, ControlGetFocus("A")      ; scroll Right - 1 is SB_LINERIGHT ; same as Loop 1

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
; ahk_class XLMAIN  ; may apply to MS Excel - modified from https://superuser.com/a/825291/391770

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

Still not working? Try other niche solutions mentioned here - https://superuser.com/questions/13763/horizontal-scrolling-shortcut-in-windows
*/

;------------------------------------------------------------------------------
;  = Media Keys Restored (disabled)
; uncomment to use, if media keys are remapped to navigation keys in "Remap Keys" section

/*
#HotIf WinActive("ahk_group MediaKeys")

; restore default actions
Media_Next::Media_Next
Media_Prev::Media_Prev
Media_Stop::Media_Stop
Media_Play_Pause::Media_Play_Pause

#HotIf
*/

;------------------------------------------------------------------------------
;  = Symbols In File Names keys
; inspired by the file renaming scheme of "yt-dlg" app - https://oleksis.github.io/youtube-dl-gui/

#HotIf WinActive("ahk_group FileNameSymbols")

; replace \/:*?"<>| with ＼⧸ ： ✲ ？＂＜＞｜
; comment out the ones you don't desire, like \ → ＼
; :?*:\::{U+FF3C}                     ; \ → ＼ | replace U+005C REVERSE SOLIDUS : backslash            → U+FF3C FULLWIDTH REVERSE SOLIDUS   ; disabled

:?*:/::{U+29F8}                     ; / → ⧸  | replace U+002F SOLIDUS : slash, forward slash, virgule → U+29F8 BIG SOLIDUS
:?*:`:::{U+FF1A}                    ; : → ：  | replace U+003A COLON                                  → U+FF1A FULLWIDTH COLON
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

;  = Date & Time

;    + Format Date / Time

:*x:d++::      Send FormatTime(, "yyyy.MM.dd")          ; sends 2021.02.31
:*x:date+::    Send FormatTime(, "dd.MM.yyyy")          ; sends 28.03.2020
:*x:time+::    Send FormatTime(, "h:mm tt")             ; sends 6:48 PM
:*x:datetime+::Send FormatTime(, "dd/MM/yyyy h:mm tt")  ; sends 28/03/2020 6:46 PM

;------------------------------------------------------------------------------
;  = URL Encode/Decode

:*x:url+::Send UrlEncode(A_Clipboard)

/* Encode URL
    Example: https://www.google.com/
    Copy example URL to clipboard
    Triger `UrlEncode` function by typing `url+`
    Output: https%3A%2F%2Fwww.google.com%2F
*/

:*x:url-::Send UrlDecode(A_Clipboard)

/* Decode URL
    Example: https%3A%2F%2Fwww.google.com%2F
    Copy example URL to clipboard
    Triger `UrlDecode` function by typing `url-`
    Output: https://www.google.com/
*/

;------------------------------------------------------------------------------
;  = Find & Replace in Clipboard

;    + Find & Replace dot with space

:*:.++:: { ; hotstring ".++"
A_Clipboard := StrReplace(A_Clipboard,"."," ") ; replace each dot with space
}

/*
Find text:          "ABC..def.GHI"
Replacement text:   "ABC  def GHI"
*/

;--------
;    + Find & Replace dot with space (RegEx)

:*:.r++:: { ; hotstring ".r++"
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
cliptext := RegExReplace(A_Clipboard, "\s+", A_Space) ; replace \s with A_Space = " " (single space)
A_Clipboard := RegExReplace(cliptext, "^ +| +$")      ; trim leading/trailing spaces
}

/* Regular Expressions (RegEx) -
\s = [\r\n\t\f\v ]
\r = carriage return
\n = line-feed (newline)
\t = tab
\f = form-feed
\v = vertical white space
" "= white space
see Regular Expressions (RegEx) - Quick Reference in AHK help file for more information

Example text - Input:
Line0 ""              (blank line)
Line1 " FUBFUBFI    " (leading and trailing space)
Line2 "dvvbvvoe   df" (2+ consecutive spaces)
Line3 ""              (blank line)

Trimmed text - Output:
Line1 "FUBFUBFI dvvbvvoe df"
Explanation: blank lines are deleted, one or more \s are replaced with space, resulting in Line2 being appended to Line1
*/

;--------
; Trim but keep non-blank lines

:?*:v++:: { ; hotstring "v++"
cliptext := RegExReplace(A_Clipboard, "m)^ +| +$")   ; m) = multi-line, trim leading/trailing spaces
cliptext := RegExReplace(cliptext, "^\s+|\s+$|`r")   ; trim \r, leading/trailing \n
A_Clipboard := RegExReplace(cliptext, " +", A_Space) ; trim double spaces
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
; User-defined functions

;  = MyNotification
; search for `ToolTipFn` for alternative

;    + MyNotificationGui

MyNotificationGui(mytext, myduration := -500, xAxis := 1550, yAxis := 985, timer := 1) { ; 500ms
Global MyNotification := Gui("+AlwaysOnTop -Caption +ToolWindow")   ; +ToolWindow avoids a taskbar button and an Alt-Tab menu item.
MyNotification.BackColor := "2C2C2E"                ; "2C2C2E" for dark mode ; "EEEEEE" for White background ; can be any RGB colour (it will be made transparent below)
MyNotification.SetFont("s9 w1000", "Arial")         ; font size 9, bold
MyNotification.AddText("cWhite w230 Left", mytext)  ; "cWhite" for dark mode ; use "cBlack" for black text on white background
MyNotification.Show("x1650 y985 NoActivate")        ; NoActivate avoids deactivating the currently active window
WinMove xAxis, yAxis,,, MyNotification
If timer = 1
    SetTimer () => EndMyNotif(), myduration ; 500ms ; new thread
If timer = 0 {
    Sleep myduration * -1
    EndMyNotif()
    }
}

;--------
;    + EndMyNotif

EndMyNotif() {
MyNotification.Destroy()
}

;------------------------------------------------------------------------------
;  = Toggle protected operating system (OS) files
; inspiration from https://www.autohotkey.com/board/topic/82603-toggle-hidden-files-system-files-and-file-extensions/?p=670182

;    + ToggleOS

ToggleOS(*) {
; alternative - Run ToggleSystemFiles.bat as administrator to toggle settings - https://superuser.com/a/1151851/391770
If Status = 0 { ; enable if disabled
    RegWrite "1", "REG_DWORD", "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSuperHidden"
    CheckRegWrite(Status)
    ToggleOSCheck
    SetTimer () => WindowsRefreshOrRun(), -100       ; 100ms ; new thread
    }
Else { ; disable if enabled
    RegWrite "0", "REG_DWORD", "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSuperHidden"
    CheckRegWrite(Status)
    ToggleOSCheck
    SetTimer () => WindowsRefreshOrRun(), -100       ; 100ms ; new thread
    }
}

;--------
;    + CheckRegWrite

CheckRegWrite(value) { ; check if RegWrite was success
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
Sleep 2000                                  ; 2s wait for registry change to be enforced
If WinExist("ahk_class CabinetWClass") {    ; If Windows File Explorer window exists
    WinActivate
    If WinWaitActive(,, 2)                  ; 2s timeout ; wait for explorer to become active window
        RefreshExplorer()
    }
Else { ; open new explorer window if one doesn't already exist ; comment out this section if not desired
    Run 'explorer.exe',,"Max"
    WinWait("ahk_class CabinetWClass",, 10) ; 10s timeout
    WinActivate
    }
}

;--------
;    + RefreshExplorer
; Source: https://www.autohotkey.com/boards/viewtopic.php?p=482766#p482766

RefreshExplorer() {
    Local Windows := ComObject("Shell.Application").Windows
    Windows.Item(ComValue(0x13, 8)).Refresh()
    For Window in Windows
        If (Window.Name != "Internet Explorer")
            Window.Refresh()
}

/* alternatives - but not tested personally, and likely not as reliable (check source)

N°. 1 -
Sleep 2000 ; 500 ms
Send "{F5}"             ; refresh
; a second refresh might be needed after a few seconds to see the effects of change in settings
; add a Sleep command or use SetTimer prior to refresh to account for the delay

N°. 2 - source: https://www.autohotkey.com/boards/viewtopic.php?p=543680#p543680
ExplorerRefresh() => PostMessage(WM_COMMAND := 0x111, Refresh := 41504,, HWND_BROADCAST := 0xFFFF)

N°. 3 - source: https://www.autohotkey.com/boards/viewtopic.php?p=543680#p543680
UpdateWindows()
{
    _ttm := A_TitleMatchMode
    SetTitleMatchMode 'RegEx'
    for window in WinGetList("ahk_class ExploreWClass|CabinetWClass|Progman")
        PostMessage 0x111 , 41504 ,,, window
    SetTitleMatchMode _ttm
}

N°. 4 - Source: https://www.autohotkey.com/board/topic/12342-showhide-hidden-files-and-folders/#entry79944
PostMessage 0x111, 41504
*/

;------------------------------------------------------------------------------
;  = Launch explorer or reuse to open path

;    + OpenFolder

OpenFolder(path) {
If WinExist("ahk_class CabinetWClass") {           ; if explorer is open
    WinActivate
    If WinWaitActive(,, 2) { ; 2s = Sleep 2000, but sends next command as soon as activated, instead of waiting for the full 2000ms period
    
        ; wait for cursor focus
        If FocusExplorerAddressBar() == "err0r" {
            Run 'explore "' path '"',,"Max"
            Exit
            }
        
        ; check to see if existing path is not equal to new path
        Else If path !== CallClipboardVar(2, 1) { ; 2s, Return
            Send "{Raw}" path
            WinWaitClose(,, 2) ; 2s - wait for drop down to disappear, then Send Enter ; WinWait commands used to prevent drop down display appearing after Enter - explorer bug
            Send "{Enter}{F6 2}"
            }
            
        ; if paths are equal, no further change is necessary, refocus on file list
        Else Send "{F6 2}"
        }
    }
; if explorer is not open
Else Run 'explore "' path '"',,"Max"
}

;--------
;    + FocusExplorerAddressBar
; check focus of cursor in explorer before further action

FocusExplorerAddressBar() {
Send "{F4}"
While ControlGetClassNN(ControlGetFocus("A")) !== "Microsoft.UI.Content.DesktopChildSiteBridge1" {
    Sleep 100
    If A_Index > 5 { ; = Sleep 500 ; wait until focus is on address bar, max 500ms
        ToolTipFn("Failed to focus address bar", -1000) ; 1s
        Return "err0r"
        }
    }
If WinWait("PopupHost ahk_class Microsoft.UI.Content.PopupWindowSiteBridge",, 2) ; 2s - wait for drop down
    Return "Success"
Else Return "err0r"
}

;--------
;  = Adjust Window Transparency

;    + GetTrans

GetTrans(id) {
Trans := WinGetTransparent("ahk_id " id)
If not Trans
    Trans := 255
Return Trans
}

;--------
;    + SetTransByWheel

SetTransByWheel(Transparency, id) {
If Transparency = "Off"
    WinSetTransparent 255, "ahk_id " id
    ; Set transparency to 255 before using Off - might avoid window redrawing problems such as a black background. If the window still fails to be redrawn correctly, try WinRedraw, WinMove or WinHide + WinShow for a possible workaround.
WinSetTransparent Transparency, "ahk_id " id
ToolTipFn("Transparency: " Transparency) ; 500ms
}

;--------
;    + SetTransMenuFn
; modified from http://www.computoredge.com/AutoHotkey/Downloads/Always_on_Top.ahk

SetTransMenuFn() {
MouseGetPos ,, &WinID   ; identify window id
; WinID := WinExist("A")  ; alternative - but 'Active' window might not always be the intended target
Global WinID            ; so that SetTransByMenu can use it to set transparency
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
Transparency := Trim(SubStr(item, 4, 3))
WinSetTransparent Transparency, "ahk_id " WinID
If Transparency = 255 {
    WinSetTransparent "Off", "ahk_id " WinID ; Specifying Off - may improve performance and reduce usage of system resources
    }
ToolTipFn("Transparency: " Trim(SubStr(item, 4)), -2000) ; 2s
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
CallClipboard(2) ; 2s, Exit
CaseConvert(StrLower(A_Clipboard))
}

;--------
;    + ConvertSentence

ConvertSentence(*) {
CallClipboard(2) ; 2s, Exit
CaseConvert(RegExReplace(StrLower(A_Clipboard), "(((^\s*|([.!?]+\s*))[a-z])|\Wi\W)", "$U1")) ; Code Credit #1
}

; RegEx explanation -
; \s = [\r\n\t\f\v ]
; $U1 = back reference uppercase 1
; \W = [^a-zA-Z0-9_] = any character that is NOT alphabet, number, underscore

;--------
;    + ConvertTitle

ConvertTitle(*) {
CallClipboard(2) ; 2s, Exit
CaseConvert(StrTitle(A_Clipboard))
}

;--------
;    + ConvertUpper

ConvertUpper(*) {
CallClipboard(2) ; 2s, Exit
CaseConvert(StrUpper(A_Clipboard))
}

;--------
;    + ConvertInvert

ConvertInvert(*) {
CallClipboard(2) ; 2s, Exit
inverted := ""
Loop Parse A_Clipboard {     ; Code Credit #2
    If StrLower(A_LoopField) == A_LoopField    ; * Code Credit #3
        inverted .= StrUpper(A_LoopField)      ; *
    Else inverted .= StrLower(A_LoopField)     ; *
    }
CaseConvert(inverted)
}
; Unicode TestString  :=
; abcdefghijklmnopqrstuvwxyzéâäàåçêëèïîìæôöòûùÿáíóúñABCDEFGHIJKLMNOPQRSTUVWXYZÉÂÄÀÅÇÊËÈÏÎÌÆÔÖÒÛÙŸÁÍÓÚÑ
; Unicode iNVERT cASE :=
; ABCDEFGHIJKLMNOPQRSTUVWXYZÉÂÄÀÅÇÊËÈÏÎÌÆÔÖÒÛÙŸÁÍÓÚÑabcdefghijklmnopqrstuvwxyzéâäàåçêëèïîìæôöòûùÿáíóúñ

;--------
;    + CaseConvert

CaseConvert(caseText) {
string := StrReplace(caseText, "`r") ; remove \r
Len := StrLen(string)
A_Clipboard := string
Send "^v"    ; Paste text
If Len <= 20 ; and select text only if text ≤ 20 characters (change limit as needed)
    Send "+{Left " Len "}"
}

; Code Credit #1 - NeedleRegEx pattern modified from https://www.autohotkey.com/board/topic/24431-convert-text-uppercase-lowercase-capitalized-or-inverted/?p=158295
; Code Credit #2 - idea for loop modified from https://www.autohotkey.com/boards/viewtopic.php?p=58417#p58417
; Code Credit #3 - 3 lines of code with a comment "; *" were adapted from a (inaccurate) answer generated from a query to KudoAI's DuckDuckGPT user script - https://greasyfork.org/en/scripts/459849-duckduckgpt

;------------------------------------------------------------------------------
;  = Clipboard functions

;    + CallClipWait

CallClipWait(secs) {
ToolTipFn("Waiting for clipboard", secs * -1000) ; 2s
If not ClipWait(secs) {
    ToolTipFn(A_ThisHotkey ":: Clip Failed", -2000) ; 2s
    ; MyNotificationGui(A_ThisHotkey ":: Clip Failed", -2000) ; 2s ; alternative to tooltip
    Exit
    }
}

;--------
;    + CallClipboard

CallClipboard(secs, retrn := 0) {
Global clipSave := ClipboardAll() ; Global = Return clipSave
A_Clipboard := ""
Send "^c"
ToolTipFn("Waiting for clipboard", secs * -1000) ; 2s
If not ClipWait(secs) {
    ToolTipFn(A_ThisHotkey ":: Clip Failed", -2000) ; 2s
    A_Clipboard := clipSave
    If retrn = 0
        Exit
    Else Return "err0r" ; If retrn = 1
    }
}

;--------
;    + CallClipboardVar

CallClipboardVar(secs, retrn := 0) {        ; copied text is sent to variable, clipboard is restored
If CallClipboard(secs, retrn) == "err0r"    ; ClipChanged is turned on
    Return "err0r"
Else {
    clipped := A_Clipboard
    A_Clipboard := clipSave
    Return clipped
    }
}

;------------------------------------------------------------------------------
;  = ToolTip functions

;    + ToolTipFn

ToolTipFn(mytext, myduration := -500, xAxis?, yAxis?) { ; 500ms

Try ToolTip(,,, WhichToolTip) ; turn Off previous ToolTip

; set/increment WhichToolTip parameter
static WhichToolTip := 1
If WhichToolTip = 20
    WhichToolTip := 1
Else WhichToolTip++

ToolTip mytext, xAxis?, yAxis?, WhichToolTip
SetTimer () => ToolTip(,,, WhichToolTip), myduration ; 500ms ; new thread
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
WrapTextMenu.Add("&3   (  Round Brackets )"         ,WrapTextFromMenu) ; round brackets ()
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
    EncText("(",")")        ; enclose in round brackets ()
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
CallClipboard(2) ; 2s, Exit
TextString := StrReplace(A_Clipboard, "`r")             ; remove \r for StrLen
TextStringInitial := TextString                         ; save initial string for later
TextString := RegExReplace(TextString,'^\s+|\s+$')      ; RegEx remove leading/trailing ; \s = [\r\n\t\f\v ]
TextString := RegExReplace(TextString,'^[\[`'\(\{%`"“‘]+|^``',, &ReplacementCount)     ;"; remove leading  ['({%"“‘`  ; customise as your needs in WrapTextMenuFn and WrapText Keys
If ReplacementCount > 0 ; don't remove trailing character if leading character is not removed
    TextString := RegExReplace(TextString,'[\]`'\)\}%`"”’]+$|``$')     ;"; remove trailing ]')}%"”’`  ; customise as your needs in WrapTextMenuFn and WrapText Keys
TextString := q TextString p
Len := StrLen(TextString)

; If you regularly include leading/trailing spaces within quotes, comment out above RegEx and below If statements
If RegExMatch(TextStringInitial, "^\s+", &Lead) {   ; If the initial string has leading \s
    TextString := Lead[] TextString  ; add &OutputVar to string
    Len += Lead.Len                  ; add the length of &OutputVar to Len
    }
If RegExMatch(TextStringInitial, "\s+$", &Trail) {   ; If the initial string has trailing \s
    TextString .= Trail[]            ; append &OutputVar to string
    Len += Trail.Len                 ; add the length of &OutputVar to Len
    }

; Send "{Raw}" TextString ; send the string with quotes
A_Clipboard := TextString ; pasting from clipboard is faster than send raw, especially for long strings
Send "^v"                 ; paste
If Len <= 20              ; and select text string if ≤ 20 characters (change limit as needed)
    Send "+{Left " Len "}"
; A_Clipboard := TextStringInitial  ; restore original text string to clipboard if desired
}

;------------------------------------------------------------------------------
;  = URL Encode/Decode
; Modified from https://www.autohotkey.com/boards/viewtopic.php?style=7&t=116056#p517193

;    + UrlDecode

UrlDecode(Url, Enc := "UTF-8") {
; Validate url
; If not RegExMatch(Url, "^http(s|)%3A%2F%2F[\w-@%~#&=\.\+\?]{2,256}\.[a-z]{2,6}($|%2F[\w-@%~#&=\.\+\?]*$)")
;     Return
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
ToolTipFn("Decoding successful!", -2000) ; 2s
Return Url
}

;--------
;    + UrlEncode

UrlEncode(str, sExcepts := "-_.", enc := "UTF-8") {
validateURL := "^http(s|)://[\w-@:%~#&=\.\+\?]{2,256}\.[a-z]{2,6}($|/[\w-@:%~#&=/\.\+\?]*$)" ; modified from https://www.makeuseof.com/regular-expressions-validate-url/
If not RegExMatch(str, validateURL) {
    Try UrlDecode(str) ; url is already encoded
    Catch {
        MsgBox "ERROR!`n`nString: " str "`nRegEx: " validateURL "`n`nNot a valid URL.",, "262144"
        Exit
        }
    Else ToolTipFn("URL already encoded!", -2000) ; 2s
    Return str
    }
Else {
    hex := "00"
    buff := Buffer(StrPut(str, enc))
    StrPut(str, buff, enc)
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
            DllCall("msvcrt\swprintf", "Str", hex, "Str", "%%%02X", "UChar", b, "Cdecl")
            encoded .= hex
            }
        }
    ToolTipFn("Valid URL! Encoding successful!", -2000) ; 2s
    Return encoded
    }
}

;------------------------------------------------------------------------------
;  = Kill All Instances Of An App

;    + GetKillTitles

GetKillTitles(HWNDs) {
temp := ""
Loop HWNDs.Length {
    t := RegExReplace(WinGetTitle(HWNDs[A_Index]), "\s+", A_Space)   ; remove \s = [\r\n\t\f\v ]
    If t = "" and A_Index < 10
        tempTitle := "  " A_Index " = #No Title`n    "   ; pad space before 1-9
    Else If t = ""
        tempTitle := A_Index " = #No Title`n    "
    Else If A_Index < 10
        tempTitle := "  " A_Index " = " t "`n    "
    Else tempTitle := A_Index " = " t "`n    "

    ; rudimentary word wrap ; for more sophisticated solution (AHK v1) - https://www.autohotkey.com/boards/viewtopic.php?t=59461
    If StrLen(tempTitle) > 61
        temp .= SubStr(tempTitle, 1, 50) "-`n            -" SubStr(tempTitle, 51, 50)
    Else temp .= tempTitle
    }
Return temp
}

;------------------------------------------------------------------------------
;  = Print Screen

;    + SnipMenuFn

SnipMenuFn() {
SnipMenu := Menu()
SnipMenu.Delete
SnipMenu.Add("&1 Rectangular Snip"          ,SnipFromMenu)
SnipMenu.Add("&2 Window Snip"               ,SnipFromMenu)
SnipMenu.Add("&3 Full Screen Snip"          ,SnipFromMenu)
SnipMenu.Add("&4 Freeform Snip"             ,SnipFromMenu)
SnipMenu.Show
}

;--------
;    + SnipFromMenu
; Applicable to SnippingTool.exe version 11.2407.3.0 and later (Date: 2024.09.16)

SnipFromMenu(ItemName, ItemPos, MyMenu) {
Send "{PrintScreen}"
If WinWaitActive("Snipping Tool Overlay ahk_exe SnippingTool.exe",, 3) { ; 3s
    Sleep 250
    MouseMove A_ScreenWidth/2, 40 ; client 973, 38
    MouseClick
    If WinWait("PopupHost ahk_exe SnippingTool.exe",, 3) { ; 3s
        Sleep 250
        Send "{Up 4}"
        Send "{Down " ItemPos - 1 "}{Enter}"
        }
    Else {
        ToolTipFn(A_ThisHotkey ":: Snipping Tool Pop-up timed out", -2000) ; 2s
        Exit
        }
    }
Else {
    ToolTipFn(A_ThisHotkey ":: Screen Snipping Overlay timed out", -2000) ; 2s
    Exit
    }
; wait for screenshot to be taken ; abort further action if timeout or Esc key is pressed
If WinWaitClose("Snipping Tool Overlay ahk_exe SnippingTool.exe",, 15) and (A_PriorKey != "Escape") { ; 15s

    ; If `Automatically save screenshots` is ENABLED in snippingtool
    MyPath := "C:\Users\" A_UserName "\Pictures\Screenshots\Screenshot " FormatTime(, "yyyy-MM-dd HHmmss") ".png"
    SetTimer () => ScreenshotFileOp(MyPath), -100 ; 100ms ; new thread
    
/*  ; If `Automatically save screenshots` is disabled in snippingtool, use below code to open paint and edit/save from clipboard

    If WinExist("ahk_class MSPaintApp")
        WinActivate
    Else Run "mspaint.exe",,"Max"

    If WinWait("ahk_class MSPaintApp",, 3) { ; 3s
        Sleep 500 ; 500ms
        PostMessage 0x111, 57637,,, "ahk_class MSPaintApp"  ; 0x111 = Paste
        ; Send "^v"                                         ; alternative to PostMessage
        ; ControlSend "^v",, "ahk_class MSPaintApp"         ; alternative to Send / PostMessage
        PostMessage 0x111, 620,,, "ahk_class MSPaintApp"    ; activate "Select" tool
        }
 */
    }
Else If A_PriorKey = "Escape"
    ToolTipFn(A_ThisHotkey ":: Screen Snipping aborted - Esc", -2000) ; 2s
Else ToolTipFn(A_ThisHotkey ":: Screen Snipping aborted - 15s timeout", -2000) ; 2s
}

/* For older versions of screen snippingtool, below code may work

SnipFromMenu(ItemName, ItemPos, MyMenu) {
Send "{PrintScreen}"

; wait for screenshot tool to activate
If WinWaitActive("Screen Snipping ahk_class Windows.UI.Core.CoreWindow",, 3) { ; 3s timeout
    Sleep 500
    Send "{Tab " ItemPos "}{Enter}"
    }
Else {
    ToolTipFn(A_ThisHotkey ":: Screen Snipping timed out", -2000) ; 2s
    Exit
    }

; wait for screenshot to be taken ; abort further action if timeout or Esc key is pressed
If WinWaitClose(,, 15) and (A_PriorKey != "Escape") { ; 15s
    If WinExist("ahk_class MSPaintApp")
        WinActivate
    Else Run "mspaint.exe",,"Max"

    If WinWait("ahk_class MSPaintApp",, 3) ; 3s
        ; modified from https://www.autohotkey.com/boards/viewtopic.php?p=360163#p360163
        PostMessage 0x111, 57637,,, "ahk_class MSPaintApp" ; 0x111 = Paste
        ; Send "^v"                                 ; alternative to PostMessage
        ; ControlSend "^v",, "ahk_class MSPaintApp" ; alternative to Send / PostMessage
    }
Else ToolTipFn(A_ThisHotkey ":: Screen Snipping aborted - 15s timeout / Esc", -2000) ; 2s
}

;--------
; use PostMessage to select tool -
Free-Form       621
Select          620
Eraser          637
Fill            623
Pick            639
Magnifier       638
Pencil          636
Brush           640
AirBrush        627
Text            622
Line            624
Curve           628
Rectangle       641
Polygon         632
Ellipse         643
Rounded_Rect    634

PostMessage 0x111, 620,,, "ahk_pid " <Process ID> ; alternative to "ahk_class MSPaintApp"
paint settings in registry = HKCU\Software\Microsoft\Windows\CurrentVersion\Applets\Paint\
*/

;--------
;    + PrintScreenFn

PrintScreenFn(*) {
; save screenshot number in variable 'serial'
serial := RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer", "ScreenshotIndex", "1")

; send Windows + Print Screen
Send "#{PrintScreen}"

; use SetTimer to wait for file creation 100ms and execute action in another thread
; pass serial as number for path of screenshot file in `ScreenshotFileOp` function
SetTimer () => ScreenshotFileOp("C:\Users\" A_UserName "\Pictures\Screenshots\Screenshot (" serial ").png"), -100   ; 100ms ; new thread
}

;--------
;    + ScreenshotFileOp

ScreenshotFileOp(MyPath) {
While !FileExist(MyPath) { ; if file does not exist, wait using Sleep and repeat check
    Sleep 500        ; wait for 500ms
    If A_Index > 6 { ; max 6 attempts, total wait 3000ms = 3s
        ; If failed ×6, open screenshot folder in explorer
        OpenFolder("C:\Users\" A_UserName "\Pictures\Screenshots\")
        MsgBox MyPath,, "262144"    ; show calculated path in pop-up to compare with actual file path
        Exit                        ; give up on rename
        }
    }

; prepare to rename file
fileTime := FormatTime(FileGetTime(MyPath, "C"), "yyyy-MM-dd @ HH：mm：ss") ; 2016-07-21 @ 13：28：05
; FileGetTime - obtain creation time as a string in YYYYMMDDHH24MISS format
; FormatTime - transform Timestamp YYYYMMDDHH24MISS into desired date/time format.
NewPath := "C:\Users\" A_UserName "\Pictures\Screenshots\" fileTime ".png"

FileMove MyPath, NewPath                ; rename

; Further actions - (uncomment below lines to execute)
; Run 'mspaint.exe "' NewPath '"',,"Max"                                ; open in paint
; OpenFolder("C:\Users\" A_UserName "\Pictures\Screenshots\")           ; open screenshot folder in explorer
}

;------------------------------------------------------------------------------
;  = Windows File Explorer Fn

;    + GetFolderSize

GetFolderSize(pathType, pathContent) {
; variables
FolderSizeB := 0
errorDetails := ""
folderName := ""
fileName := ""

If pathType = 1 {   ; multiple lines ; InStr(pathContent, "`n")
    Loop Parse pathContent, "`n", "`r" {
        If DirExist(A_LoopField) {          ; is folder
            Loop Files A_LoopField "\*.*", "R"
                FolderSizeB += A_LoopFileSize
            folderName .= A_LoopField "`n"
            }
        Else If FileExist(A_LoopField) {    ; is file
            FolderSizeB += FileGetSize(A_LoopField, "B")
            fileName .= A_LoopField "`n"
            }
        Else errorDetails .= "#" A_Index ": " A_LoopField " -> Not a folder or file`n"
        }

    ; display folders first in pathContent
    If folderName = ""
        pathContent := Trim(Sort(fileName), "`n")
    Else If fileName = ""
        pathContent := Trim(Sort(folderName), "`n")
    Else pathContent := Trim(Sort(folderName), "`n") "`n" Trim(Sort(fileName), "`n")

    ; limit string to 1000 characters
    If StrLen(pathContent) > 1500 {
        RegExMatch(pathContent, "[\s\S]{1,1500}`n", &output)
        pathContent := output.0 "... and more"
        }
    }
Else If pathType = 2 { ;          ; single line - folder ; DirExist(pathContent)
    Loop Files pathContent "\*.*", "R"
        FolderSizeB += A_LoopFileSize
    }
Else FolderSizeB += FileGetSize(pathContent, "B")    ; single line - file ; FileExist(pathContent)

; check size and display tooltip
If FolderSizeB = 0 and errorDetails != ""   ; size zero and error present
    ToolTipFn(errorDetails, -5000)          ; 5s ; show error
Else If FolderSizeB = 0                     ; size zero and no errors
    ToolTipFn(pathContent "`nEmpty Folder/File", -3000) ; 3s
Else ToolTipFn(pathContent . FolderSizeFn(FolderSizeB) "`n`n" errorDetails, -3000) ; 3s
}

;--------
;    + FolderSizeFn

FolderSizeFn(FolderSizeB) {
If FolderSizeB >= 1024 {
    FolderSizeKB := FolderSizeB / 1024
    If FolderSizeKB >= 1024 {
        FolderSizeMB := FolderSizeKB / 1024
        If FolderSizeMB >= 1024
            Return "`nSize: " Round(FolderSizeMB / 1024, 2) " GB"
        Else Return "`nSize: " Round(FolderSizeMB, 2) " MB"
        }
    Else Return "`nSize: " Round(FolderSizeKB, 2) " KB"
    }
Else Return "`nSize: " FolderSizeB " bytes"
}

;--------
;    + ValidPath

ValidPath(errorTxt := "") {
clipped := CallClipboardVar(2, 1) ; 2s, Return
If clipped == "err0r" {
    ToolTipFn(A_ThisHotkey ":: Error - Folder/file path copy failed!" errorTxt, -2000) ; 2s
    Exit
    }
If InStr(clipped, "`n") ; If multiple folder/files or single folder or file
    Return [1, clipped]
Else If DirExist(clipped)
    Return [2, clipped]
Else If FileExist(clipped)
    Return [3, clipped]
Else {
    MsgBox clipped "`nFolder/File does not exist!" errorTxt,, "262144"
    Exit
    }
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
ControlPanelMenu.Add("&0 Task Scheduler"                ,ControlPanelSelect) ; add more triggers using alphabets (abc…) or symbols (/*-+…)
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
    ComObject("shell.application").ControlPanelItem("taskschd.msc") ; Task Scheduler
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
ComObject("shell.application").ControlPanelItem("taskschd.msc")    ; Task scheduler

; CLSID List (Windows Class Identifiers)
::{21EC2020-3AEA-1069-A2DD-08002B30309D}                           ; Control Panel (view: small icons) ; alt
::{26EE0668-A00A-44D7-9371-BEB064C98683}                           ; Control Panel (view: category) ; alternate
::{20D04FE0-3AEA-1069-A2D8-08002B30309D}                           ; This PC
... Find more in AutoHotkey.chm::/docs/misc/CLSID-List.htm

; Others
ComObject("shell.application").ControlPanelItem("winver")          ; Windows version
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
ComObject("shell.application").ControlPanelItem("write")           ; WordPad ; N/A
ComObject("shell.application").ShutdownWindows()                   ; Shutdown Menu

System Configuration Tools (skipped items already in #x or listed above)
; Replace the Windows directory "C:\Windows" with A_WinDir as required

C:\WINDOWS\System32\UserAccountControlSettings.exe                 ; Change User Account Control Settings
C:\WINDOWS\System32\control.exe /name Microsoft.Troubleshooting    ; Modern Settings App > System > Troubleshoot
C:\WINDOWS\System32\control.exe system                             ; Modern Settings App > System > About
A_ComSpec /k %windir%\system32\ipconfig.exe                        ; View and configure network address settings
C:\WINDOWS\System32\taskmgr.exe /7                                 ; View details about programs and processes running on your computer
C:\WINDOWS\System32\cmd.exe                                        ; Open a command prompt window ; A_ComSpec
C:\WINDOWS\System32\msra.exe                                       ; Receive help from  (or offer help to) a friend over the Internet
C:\WINDOWS\System32\rstrui.exe                                     ; Restore your computer system to an earlier state
C:\WINDOWS\System32\regedt32.exe                                   ; Make changes to the Windows registry
C:\WINDOWS\System32\resmon.exe                                     ; Monitor the performance and resource usage of the local computer

Find more shortcuts to various sections within modern Settings app - https://winaero.com/ms-settings-commands-in-windows-10/
Shell:AppsFolder ; shortcuts to all apps in start menu
PostMessage 0x0111, 65305,,, "C:\YourScript.ahk ahk_class AutoHotkey" ; Suspend, Toggle
PostMessage 0x0111, 65306,,, "ScriptFileName.ahk - AutoHotkey" ; Pause, Toggle
PostMessage 0x0111, 65303,,, "ScriptFileName.ahk - AutoHotkey"  ; Reload.

restart your video drivers by pressing the key combination Win + Ctrl + Shift + B -- https://www.makeuseof.com/tag/hidden-key-combo-frozen-computer/
Classic Control Panel = control.exe
more? Check jeeswg's Explorer tutorial - https://www.autohotkey.com/boards/viewtopic.php?p=148121#p148121
*/

; End of script code

;------------------------------------------------------------------------------
; ChangeLog

/*
v2.00 - 2024.01.27
 * add changelog
 * add variable `AHKname` to easily update script name and version in template and standalone scripts
 * improve comments

v2.01 - 2024.01.28
 * rename MyNotificationFunc to MyNotificationGui
 * specify "Save As" WinTitle in `FileNameSymbols` group to avoid conflict with apps that use ahk_class #32770 for other uses, example - Notepad++ find-replace dialogue box
 * remove unnecessary parentheses in If commands
 * add exclusion for Notification Center and System tray overflow window for WinClose command
 * change WinKill command to use active window and add alternative
 * improve continuation section in ^!+F4 hotkey
 * rename GetTitles to GetKillTitles for more specificity, and move to user-defined functions
 * improve GetKillTitles - add padding, RegExreplace for \t \r \n, add word wrap
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
 * utilise 'ch' variable in UrlEncode
 * some minor changes
 * improve comments and update headings

v2.02 - 2024.01.29
 * improve ListLines WinWait command by using variables

v2.03 - 2024.01.30
 * rename function names with `Func` in the name to `Fn` because `Func` is a class
 * fix Toggle Window On Top - change WinSetTitle command to apply to known variable `t` instead of "A"
 * other minor changes

v2.04 - 2024.01.31
 * improve remap keys section to show more variations of key names, symbols and formatting
 * replace `HKEY_CURRENT_USER` with `HKCU` in reg keys
 * rename ShowSuperHidden_Status to Status for future expansion of function
 * condense ToggleOSCheck function
 * improve comments and update headings

v2.05 - 2024.02.04
 ★ add 'Print Screen' section
 - remove Telegram from 'CloseWithQW' group - conflict with default behaviour of 'Esc'
 + add 'MediaInfo in mpc' to 'CloseWithQW' group
 + add alternative 'PostMessage' to 'WinMinimize'
 + add alternative 'WinExist' to 'MouseGetPos'
 * fix position of parentheses in 'WinClose' with '!RButton'
 * fix alternative `^!F4` hotkey - add missing MouseGetPoI
 - remove unnecessary variable Process_Name from alternative `^!F4` hotkey
 * improve 'Adjust Window Transparency keys' - call 'MouseGetPos' once instead of twice for each mouse key
 * fix 'Adjust Window Transparency keys' - 'If Trans' statements in `^+WheelDown` hotkey
 * change 'A_ThisHotkey' to 'ThisHotkey' when applicable for more reliability
 * improve 'Recycle Bin shortcut' - add 'WinWait' to prevent dropdown explorer bug
 - remove unnecessary variable 'SwappedLetters'
 * fix colon replacement with U+003A in 'Symbols In File Names keys'
 * fix 'CheckRegWrite' - uncomment 'Exit' command, to stop further execution on 'RegWrite' failure
 - remove unnecessary 'ToolTip' command from 'GetTrans'
 * improve 'SetTransByWheel' - add 'WinSetTransparent' 255 before setting "Off", and replace 'ToolTip'/'SetTimer' combo with 'ToolTipFn' and place it after 'WinSetTransparent' command
 * fix 'SetTransMenuFn' - get 'WinID' before executing 'SetTransByMenu'
 * improve 'SetTransByMenu' - add 'WinSetTransparent' "Off" if 255, and add ToolTipFn
 * improve comments and update headings

v2.06 - 2024.02.05
 + add defaults to 'MyNotificationGui' parameters
 - remove default values from all 'MyNotificationGui' func calls
 + add defaults to 'ToolTipFn' parameters
 - remove default values from all 'ToolTipFn' func calls
 - remove unnecessary quotation marks "" for 'MyNotificationGui' and 'ToolTipFn' parameters
 - remove unnecessary Space or title from 'WinWait' commands
 * replace 'RegExReplace' with 'StrReplace' where possible to improve performance
 - remove unnecessary 'ToggleOSCheck' from 'ToggleOS(*)'
 * improve 'CheckRegWrite' and 'ToggleOSCheck' - call 'RegRead' only once per 'RegWrite'
 ? improve 'Recycle Bin shortcut' - replace 'WinWait' with {F6 2}
 * fix 'CaseConvert' - remove \r from 'caseText' before assigning to 'A_Clipboard'
 * fix 'EncText' - remove \r from 'A_Clipboard' before assigning to 'TextString' and 'TextStringInitial'; instead of single Space, use '&OutputVar' to modify 'TextString' and 'Len'
 * improve 'EncText' - call 'A_Clipboard' only once, rename variable 'Len1' to 'Len'
 - remove unnecessary variable 'Len2' from 'EncText'
 * improve 'CaseConvert' and 'EncText' - select text only if string is ≤ 20 characters (change limit as needed), this is to prevent sending large number of keystrokes when these functions are used for big chunks of text
 - remove unnecessary 'func' variable from 'UrlEncode'
 * improve comments
 * improve changelog - use "fix" instead of "correct/update", use "+" for new additions and "-" for removals, "★" for new functions/sections instead of "*"

v2.07 - 2024.02.20
 * improve comments

v2.08 - 2024.03.15
 ★ add 'OpenFolder' function and run 'Recycle Bin shortcut' through it
 ★ add 'CallClipboardVar' function and improve 'Exchange adjacent letters' function by using it, instead of calling clipboard multiple times
 + add disabled 'Media Keys Group' and 'Media Keys Restored' sections
 + add 'ahk_class #32770' dialogue box to 'FileNameSymbols' group
 + add 'NewThread' function and launch 'CapsWait' function through it
 * change position of "CapsLock" notification, to be more centred
 * change "!=" to "!==" wherever applicable to enable case sensitivity
 * change "&WinID" to "&id" in "Adjust Window Transparency keys" section, in order to differentiate from Global variable 'WinID' used by "SetTransMenuFn"
 + add "^+f" shortcut to close find bar in Firefox
 + add "^+h" shortcut to go Home in Firefox
 + add "^i" 'Invert selection' to  Windows File Explorer section
 + add "^!s" 'Show folder size in ToolTip' to Windows File Explorer section
 - remove unnecessary clipboard call from !n and ^!n 'Copy file names' hotkeys
 * change .r+ hotstring to .r++ in "Find & Replace dot with space (RegEx)" section
 * improve RegEx needles and optimise in "= Find & Replace in Clipboard" section
 * change MyNotificationGui colour scheme to white text on dark background (dark mode)
 * improve "WindowsRefreshOrRun" by adding 2s Sleep and launch using "NewThread"
 * change < 21 to <= 20 wherever applicable
 * improve "ToolTipFn" by adding 'xAxis?, yAxis?' optional parameters
 * replace "Windows version" with Task scheduler in 'Control Panel Tools'
 * update headings, spelling
 * improve comments and small changes
 * change changelog order for easier access

v2.09 - 2024.09.16
 ★ add `RefreshExplorer` function to improve `ToggleOSCheck` -- closes Issue #1 (Yipee! My first issue AND closure!)
 + add `Ctrl + G` UnGroup shortcut to `Windows File Explorer` section
 * change `myduration` argument in `MyNotificationGui` function to use negative numbers because negative Sleep is smaller error than forever cycling SetTimer AND to match ToolTipFn; consequently switch negative multiplier from SetTimer to Sleep 
 + add `Renaming ahk_exe qbittorrent.exe` to `FileNameSymbols` group
 + add `NoWrapText` group to replace the lone `WinActive` exclusion for `WrapTextMenuFn` function
 * update SetTimer for more versatility in `End auto-execute` section and `MyNotificationGui` function
 * improve `OpenFolder` by moving/changing some commands to new `FocusExplorerAddressBar` function
 * improve `OpenFolder` by adding a new check to see if existing path is not equal to new path; If equal, no further change is necessary, refocus on file list
 * change default notification from `MyNotificationGui` to `ToolTipFn` wherever appropriate -- personal preference shouldn't affect a public script
 * improve `CallClipboard` and `CallClipboardVar` to return error messages when required; and add `ToolTipFn` notification while waiting for clipboard
 * improve `ToolTipFn` to allow for position and add numbering
 - remove `NewThread` function because it is redundant and confusing; replace with `SetTimer`
 + add RegEx url validation and notifications to `UrlEncode` and `UrlDecode`
 * fix `SnipFromMenu` for new version of screenshot v11.2407.3.0 and later (as of 2024.09.16). Old code is still present under block comment
 * rename `PrintScreenExec`to `ScreenshotFileOp` - to reflect the file operations performed by the function and rename variables; add notification on failure
 * update `SnipMenuFn` as per new order of options in snipping tool, open automatically saved file in mspaint 
 * improve `GetFolderSize` by moving/changing some commands to new `ValidPath` function and improve ToolTips
 + add new shortcut for `Display Off` section and update existing `Esc` shortcut to RControl specifically
 * update `Display Off` section to work with updated shortcuts
 * improve `CheckRegWrite` to use `==` case sensitive operator since registry values can be non-numeric
 * rearrange/rename some functions and update headings
 * improve comments and small changes

v2.10 - 2024.10.11
 * rename file by replacing `#` with `No-` to avoid GitHub conflict with issue numbering
 ★ add `Dark ToolTip` section to adapt `ToolTipFn` function for windows dark mode
 * improve `ToolTipFn` function by removing unnecessary commands, change variable `ToolTipNo` to `WhichToolTip` (to match AHK docs) and change it from `Global` to `static` variable
 - remove `ToolTipOff` function
 * `Process Priority` shortcut changed from 'Win + Z' to `Win + P` (formerly Project shortcut) because Win + Z is newly designated shortcut for Snap Layouts
 - remove unnecessary variable `transformed` from `ConvertSentence` function
 * fix `EncText` function from removing trailing character unintentionally if leading character is not removed and remove unnecessary `StrReplace` command (due to previous `RegExReplace`)
 * use existing `OpenFolder` function to Run explorer command in `Print Screen` section
 * change `winver` to `taskschd.msc` in `Control Panel Tools`
 * rearrange/rename some function headings and update TOC
 * improve comments and small changes

v2.11 - 2024.10.15
 * rename `Dark ToolTip` section to `AHK Dark Mode` - to include all lib scripts pertaining to dark mode AHK v2
 * change dark mode ToolTip lib file from `ToolTipOptions.ahk` to `SystemThemeAwareToolTip.ahk`
 ★ add `Dark_MsgBox.ahk` and `Dark_WindowSpy` to lib and rename/modify for easier include and tracking
 * fix `AHKname` version increment
 * improve comments
*/
