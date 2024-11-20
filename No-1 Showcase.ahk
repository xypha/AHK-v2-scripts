; https://github.com/xypha/AHK-v2-scripts/edit/main/No-1%20Showcase.ahk
; Last updated 2024.11.20

; Visit AutoHotkey (AHK) version 2 (v2) help for information - https://www.autohotkey.com/docs/v2/
; Search for below commands/functions by using control + F on the help webpage - https://www.autohotkey.com/docs/v2/lib/

; comments begin with semi-colon ";" at start of line or space+; " ;" in middle of line
; comments can also be show like this - "/*" comment text "*/"
; and these two methods can be combined too :)

    /* AHK v2 No-1 Showcase - CONTENTS */
; Settings
; Auto-execute
;  = Set default state of Lock keys
;  = AHK Dark Mode
;  = Show/Hide OS files
;  = Customise Tray Icon
;  = Capitalise first letter opt-in Group
;  = Close With Esc/Q/W Group
;  = Horizontal Scrolling Group
;  = Media Keys Restored Group (disabled)
;  = Symbols In File Names Group
;  = WrapText variables
;  = WrapText Disabled Group
;  = End auto-execute
; Hotkeys
;  = Check & Reload AHK
;  = Remap Keys
;    + Disable Keys
;    + Keyboard keys
;    + Media keys (disabled)
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
;  = AHK Main Window
;  = Calculator (classic)
;  = Firefox
;  = KeePass
;  = Dark Mode - Window Spy
;  = Windows File Explorer
;    + Explorer main window
;      * Unselect
;      * UnGroup
;      * Invert selection
;      * Show folder/file size in ToolTip
;      * Delete Empty Folder
;      * Extract from folder & delete
;      * Horizontal Scrolling
;      * Copy full path
;      * Copy file names without path
;      * Copy file names without extension and path
;    + Locate desktop background
; #HotIf Groups
;  = Capitalise the first letter of a sentence
;  = Close With Esc/Q/W keys
;  = Horizontal Scrolling
;  = Media Keys Restored (disabled)
;  = Symbols In File Names keys
; Hotstrings - Actions
;  = Date & Time
;    + Format Date / Time
;  = URL Encode/Decode
; Hotstrings - Text Replacement
;  = Find & Replace in Clipboard
;    + Find & Replace dot with space (StrReplace)
;    + Find & Replace dot with space (RegEx)
;  = Trim and change text
; User-defined functions
;  = MyNotification
;    + MyNotificationGui
;    + EndMyNotif
;  = AHK Dark Mode functions
;    + ahkDarkMenu
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
;  = ToolTip function
;    + ToolTipFn
;  = Wrap Text In Quotes or Symbols
;    + WrapTextMenuFn
;    + WrapTextMenuSelectionFn
;    + WrapTextFn
;  = URL Encode/Decode
;    + UrlDecode
;    + UrlEncode
;  = Kill All Instances Of An App
;    + GetKillTitles
;    + GetKillTitlesFileList
;  = Print Screen functions
;    + SnipMenuFn
;    + SnipFromMenu
;    + PrintScreenFn
;    + ScreenshotFileOp
;  = Check Window Focus/Control
;    + CheckControlRegEx
;  = Windows File Explorer Fn
;    + GetFolderSize
;    + RBinQuery
;    + RBinVisible
;    + RBinHidden
;    + userSID
;    + checkDrivesFn
;    + RBinDisplay
;    + RBinGenerateTxt
;    + SizeFn
;    + ValidPath
;    + DeleteEmptyFolder
;    + CaptureFolderPath
;    + GetExplorerPath
;    + WallpaperPath_v4
;    + nxtBackground
;  = Check if file exists and create/append
;    + FileCreate_Or_Append
;  = Change MsgBox button text
;    + MsgBox_Custom
;    + MsgBox_ChangeButtonText
;  = Calculator view (classic)
;    + checkCalcView
;  = Control Panel Tools
;    + ControlPanelMenuFn
;    + ControlPanelSelect
;    + List of commands
; * Test
; ChangeLog

;------------------------------------------------------------------------------
; Settings
; Start of script code

#Requires AutoHotkey v2.0
#SingleInstance force
#WinActivateForce
KeyHistory 500 ; Max 500

;------------------------------------------------------------------------------
; Auto-execute
; This section should always be at the top of your script

AHKname := "AHK v2 No-1 Showcase v2.15"

; Show notification with parameters - text; duration in milliseconds; position on screen: xAxis, yAxis; timeout by - timer (1) or sleep (0)
MyNotificationGui("Loading " AHKname, 10000, 1550, 985, 1) ; 10000ms = 10 seconds, position bottom right corner (x-axis 1550 y-axis 985) on 1920×1080 display resolution; use timer

;--------
;  = Set default state of Lock keys
; Turn on/off upon startup (one-time)

SetCapsLockState    "Off"   ; CapsLock   is off - Use SetCapsLockState "AlwaysOff" to force the key to stay off permanently, and uncomment `Persistent`
SetNumLockState     "On"    ; NumLock    is ON
SetScrollLockState  "Off"   ; ScrollLock is off

;--------
;  = AHK Dark Mode
; manually comment out below lines If dark mode is NOT enabled because "#Include cannot be executed conditionally"

; check windows registry to see If dark mode is enabled
If not RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "AppsUseLightTheme", 1) {
    ; dark mode is enabled
    ; Global isLightMode := 0                         ; store RegRead results in variable for repeated use
    ahkDarkMenu()                                   ; enable dark theme for AHK menus
    }
; Else Global isLightMode := 1                        ; dark mode is NOT enabled

; download .ahk files from the `Lib` folder in this repo
; and save at the same under a `Lib` folder at the location of your .ahk script file

#Include "Lib\Dark Mode - ToolTip.ahk"            ; 2024.10.15
#Include "Lib\Dark Mode - MsgBox.ahk"             ; 2024.10.15

; Dark Mode - Window Spy                          ; 2024.10.15
; check "MyNotification.AddText" and "MyNotification.BackColor" in `MyNotificationGui` function

;--------
;  = Show/Hide OS files

A_TrayMenu.Delete                             ; Delete standard menu
A_TrayMenu.Add "&Toggle OS files", ToggleOS   ; User-defined function
A_TrayMenu.Add                                ; Add a separator
A_TrayMenu.AddStandard                        ; Restore standard menu
ToggleOSCheck()                               ; Query registry and check/uncheck

;--------
;  = Customise Tray Icon

I_Icon := A_ScriptDir "\icons\1-512.ico"
; Icon source: https://www.iconsdb.com/caribbean-blue-icons/1-icon.html     ; CC License, see credits.md
; I like to number scripts 1, 2, 3... and link the scripts to Numpad shortcuts for easy editing -- see section on "Check & Reload AHK" below
If FileExist(I_Icon)
    TraySetIcon I_Icon

;--------
;  = Capitalise first letter opt-in Group

GroupAdd "CapitaliseFirstLetter_optIn", "ahk_class Notepad++"                       ; Notepad++ text editor

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

;--------
;  = Media Keys Restored Group (disabled)
/* ; uncomment to use, if media keys are remapped to navigation keys in "Remap Keys" section

GroupAdd "MediaKeysRestored"    , "ahk_class MediaPlayerClassicW"                   ; MPC-HC
GroupAdd "MediaKeysRestored"    , "MPC-HC D3D Fullscreen"                           ; MPC-HC Full screen
GroupAdd "MediaKeysRestored"    , "ahk_class PotPlayer64"                           ; PotPlayer
; ahk_class RegEdit_RegEdit                                                          ; Registry Editor ; by default because AutoHotkey requires UIAccess

; To enable UIAccess for scripts by default,
; Navigate to the `UiAccessCommandKeyValue` in "HKEY_CLASSES_ROOT\AutoHotkeyScript\Shell\uiAccess\Command"
; copy it and paste it as (Default) value in "HKEY_CLASSES_ROOT\AutoHotkeyScript\Shell\Open\Command"
; Source: https://blog.danskingdom.com/Prevent-Admin-apps-from-blocking-AutoHotkey-by-using-UI-Access/
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
;  = WrapText variables

Global WrapText_Leading1 := "'" , WrapText_Trailing1 := "'"      ; single quotation '' - ' U+0027 : APOSTROPHE
Global WrapText_Leading2 := '"' , WrapText_Trailing2 := '"'      ; double quotation "" - " U+0022 : QUOTATION MARK
Global WrapText_Leading3 := "(" , WrapText_Trailing3 := ")"      ; round brackets  ()
Global WrapText_Leading4 := "[" , WrapText_Trailing4 := "]"      ; square brackets []
Global WrapText_Leading5 := "{" , WrapText_Trailing5 := "}"      ; flower brackets {}
Global WrapText_Leading6 := "``" , WrapText_Trailing6 := "``"    ; accent/backtick ``
Global WrapText_Leading7 := "%" , WrapText_Trailing7 := "%"      ; percent sign %%
Global WrapText_Leading8 := "‘" , WrapText_Trailing8 := "’"      ; ‘’ - ‘ U+2018 LEFT & ’ U+2019 RIGHT SINGLE QUOTATION MARK {single turned comma & comma quotation mark}
Global WrapText_Leading9 := "“" , WrapText_Trailing9 := "”"      ; “” - “ U+201C LEFT & ” U+201D RIGHT DOUBLE QUOTATION MARK {double turned comma & comma quotation mark}

;--------
;  = WrapText Disabled Group

GroupAdd "WrapText_disabled"      , "ahk_exe mpc-hc.exe"                           ; MPC-HC

GroupAdd "WrapText_disabled"      , "ahk_class CalcFrame"                          ; Calculator (classic)
; Alt + number shortcuts are used to switch between Standard / Scientific / Programmer / Statistics calculators

; GroupAdd "WrapText_disabled"      , "ahk_class MSPaintApp"                         ; MS Paint (classic)
; commented out MSPaintApp from the group to enable wrapping text when inserting/editing text element (ClassNN: RICHEDIT50W1) - see `WrapTextFn` for more info

;--------
;  = End auto-execute

SetTimer () => EndMyNotif(), -1000  ; 1s ; new thread ; Reset notification timer to 1s after code in auto-execute section has finished running
Return                              ; Ends auto-execute

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

^!Numpad1:: {                                       ; Ctrl + Alt + Numpad1 keys pressed together
MyNotificationGui("Updating " AHKname,,, 985, 0)    ; 500ms ; use Sleep coz reload cancels timers
Reload
}

;------------------------------------------------------------------------------
;  = Remap Keys

;    + Disable Keys

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

;--------
;    + Keyboard keys

; Use Alt + Insert to toggle the 'Insert mode' and retain the key's function
; Note: ^Insert = Copy(^c) is Windows default behaviour and is not changed by this code
!Insert::Insert     ; Source: https://gist.github.com/endolith/823381

LWin & Tab::AltTab ; Left Win key works as left Alt key - disables task view

RAlt::!Space       ; Alt + Space brings up a window's title bar menu

^RCtrl::MButton    ; press Left & Right Ctrl button to simulate mouse Middle Click

RCtrl & Up::        Send "{PgUp}"  ; Page up   - use "&" to create 2-key combo shortcut
RCtrl & Down::      Send "{PgDn}"  ; Page down - use a variable number of spaces before Send command without affecting the command itself
RControl & Left::   Send "{Home}"  ; Home      - use alternate key name for RCtrl
>^Right::           Send "{End}"   ; End       - use >^ instead of Right Ctrl button and skip using "&"

!m::WinMinimize "A"         ; Alt+ M = Minimize active window
; PostMessage 0x0112, 0xF020,,, "A" ; alternative, 0x0112 = WM_SYSCOMMAND, 0xF020 = SC_MINIMIZE

/*
;--------
;    + Media keys (disabled)
; remap media keys to navigation keys
; uncomment to use and enable "MediaKeysRestored" if required

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
MyNotificationGui("CapsLock ON", 10000, 845)   ; 10000ms = 10s, change to match KeyWait timeout if needed
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

winClass := WinGetClass("ahk_id " id)                    ; Retrieves the specified window's class name
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
titleLimit := 30 ; limit the heigh of resulting message box
Process_Name := WinGetProcessName("A")
Display := "Kill all instances of this app?`n`n"    ; `n = new line
        .  "Name of process:`t" Process_Name "`n"   ; `t = tab
        .  "Path of process:`t" WinGetProcessPath("ahk_exe " Process_Name) "`n`n"
        .  "No. of visible windows: " WinGetCount("ahk_exe " Process_Name) "`n" ; no of windows ≠ no of processes
        .  "Titles of visible windows -`n"
        .  GetKillTitles(WinGetList("ahk_exe " Process_Name), titleLimit) "`n"

DetectHiddenWindows True
HWNDs := WinGetList("ahk_exe " Process_Name)
msgList := GetKillTitles(HWNDs, titleLimit)
Display .= "Total number of windows: " WinGetCount("ahk_exe " Process_Name) " (incl. hidden)`n"
        .  "Titles of all windows:`n"
DetectHiddenWindows False ; default

If HWNDs.Length <= titleLimit
    result := MsgBox(Display . msgList, A_ScriptName " - WARNING", 1 + 48 + 256 + 262144)
    ; 1 OKCancel (buttons) 48 Icon! (add Exclamation icon) 256 Default2 (make No the default button to prevent accidental process kill) 262144 = Always on top
Else result := MsgBox_Custom(Display . msgList, A_ScriptName " - WARNING", 48 + 512 + 262144, 3, "OK", "Title List", "Cancel")
    ; 3 = Yes / No / Cancel ; 48 Icon! 512 Default3 262144 = Always-on-top

If result == "OK" or result == "Yes" { ; MsgBox_Custom Button1
    While ProcessExist(Process_Name)
        ProcessClose Process_Name
    ; Run A_ComSpec ' /C Taskkill /IM /F "' Process_Name '"' ; alternative - you might see a flash of command prompt/terminal window. Brief explanation of flags -
    ; /C Carries out the command and then terminates
    ; /IM imagename
    ; /F forcefully terminates
    ; open Run dialogue (Win + R), paste "cmd.exe /?" and press OK, to see default flags
    ; then paste "Taskkill /?" (without the quotation marks) in cmd window and press enter to see 'Taskkill' specific flags, filters and examples
    }
Else If result == "No" { ; MsgBox_Custom Button2 renamed Title List
    filePath := A_Temp "\Kill list.txt"
    FileCreate_Or_Append(filePath, Display . GetKillTitlesFileList(HWNDs))
    Run '"' filePath '"',,"Max"
    }
; Else result == "Cancel" ; MsgBox_Custom Button3
;     Return
}

;------------------------------------------------------------------------------
;  = Adjust Window Transparency keys
; Modified from https://www.autohotkey.com/board/topic/667-transparent-windows/?p=148102

^+WheelUp:: {           ; increases Trans value, makes the window more opaque
MouseGetPos ,, &id
; id := WinExist("A")   ; alternative - but 'Active' window might not always be the intended target
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

F8::SetTransMenuFn()

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
Else {
    OpenFolder("::{645ff040-5081-101b-9f08-00aa002f954e}") ; comment out if not desired
    ; alternative to OpenFolder, directly open recycle bin in a new explorer window with below command
    ; Run "::{645ff040-5081-101b-9f08-00aa002f954e}"         ; If explorer is not open, then open Bin in explorer
    }

/* display number of files and size of recycle bin */
; obtain drive letters, Breakdown per drive, total no of files, total size of files, list of files
; and display results in MsgBox
RBinDisplay(RBinQuery())
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
SendMessage 0x0112, 0xF170, 2,, "Program Manager"
    ; 0x0112 is WM_SYSCOMMAND, 0xF170 is SC_MONITORPOWER.
    ; Use -1 in place of 2 to turn the monitor on.
    ; Use 1 in place of 2 to activate the monitor's low-power mode.

; further actions -
; Send "{Media_Stop}"           ; stop playing all media
; DllCall("LockWorkStation")    ; Lock Workstation - source: https://gist.github.com/raveren/bac5196d2063665d2154#file-aio-ahk-L741
}

;--------
;  = Add Control Panel Tools to a Menu

#+x::ControlPanelMenuFn()   ; Win + Shift + x

;--------
;  = Change Text Case

!c::ChangeCaseMenuFn()      ; Alt + C

;--------
;  = Wrap Text In Quotes or Symbols keys

#HotIf not WinActive("ahk_group WrapText_disabled")
; disables below hotkeys in apps that belonging to this group because they don't use it or have conflicts

!q::WrapTextMenuFn() ; Alt + Q

; WrapText Keys - Alt + Number (from the number row)
!1::WrapTextFn(WrapText_Leading1 , WrapText_Trailing1, 1)      ; enclose in single quotation '' - ' U+0027 : APOSTROPHE
!2::WrapTextFn(WrapText_Leading2 , WrapText_Trailing2, 1)      ; enclose in double quotation "" - " U+0022 : QUOTATION MARK
!3::WrapTextFn(WrapText_Leading3 , WrapText_Trailing3, 1)      ; enclose in round brackets  ()
!4::WrapTextFn(WrapText_Leading4 , WrapText_Trailing4, 1)      ; enclose in square brackets []
!5::WrapTextFn(WrapText_Leading5 , WrapText_Trailing5, 1)      ; enclose in flower brackets {}
!6::WrapTextFn(WrapText_Leading6 , WrapText_Trailing6, 1)      ; enclose in accent/backtick ``
!7::WrapTextFn(WrapText_Leading7 , WrapText_Trailing7, 1)      ; enclose in percent sign %%
!8::WrapTextFn(WrapText_Leading8 , WrapText_Trailing8, 1)      ; enclose in ‘’ - ‘ U+2018 LEFT & ’ U+2019 RIGHT SINGLE QUOTATION MARK {single turned comma & comma quotation mark}
!9::WrapTextFn(WrapText_Leading9 , WrapText_Trailing9, 1)      ; enclose in “” - “ U+201C LEFT & ” U+201D RIGHT DOUBLE QUOTATION MARK {double turned comma & comma quotation mark}
!0::WrapTextFn( ""               , ""  , 1)                    ; remove above quotes

#HotIf

;------------------------------------------------------------------------------
;  = Exchange adjacent letters
; place cursor between 2 letters. The letters reverse positions - `ab|c` becomes `ac|b`.
; Modified from http://www.computoredge.com/AutoHotkey/Downloads/LetterSwap.ahk

$!l:: { ; Alt + L ; Test: AbC
Send "{Left}+{Right 2}"
clipped := CallClipboardVar(2) ; 2s, Exit
Send SubStr(clipped,2) SubStr(clipped,1,1) "{Left}"
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
        MyNotificationGui("ERROR! Priority could not be changed!`nProcess: " Process_Name "`nPriority :  " LB.Text, 5000) ; 5s
    Else ; if successful
        MyNotificationGui("Success! Priority changed!`nProcess: " Process_Name "`nPriority :  " LB.Text, 5000) ; 5s
    Finally PPGui.Destroy()
    }
}

;------------------------------------------------------------------------------
;  = Print Screen keys

; $PrintScreen::      ; keyboard hook $ ; commented out to preserve default function
#PrintScreen:: {    ; Win + PrintScreen
PrintScreenFn       ; take screenshot, save and rename
}

; #+r::           ; video snip shortcut, uncomment if desired
^PrintScreen::  ; Ctrl + Print Screen (key name = PrtSc, PrtScn or PrntScrn)
#+s:: {         ; Win + Shift + s
SnipMenuFn
}

;------------------------------------------------------------------------------
; #HotIf Apps
; Tailor keyboard shortcuts, commands and functions to specific windows, apps or pre-defined groups of both

;  = AHK Main Window

#HotIf WinActive(".ahk - AutoHotkey v ahk_class AutoHotkey")
; because this is to enable below commands to apply on main windows of all running scrips irrespective of v1 or v2
; alternative to #HotIf WinActive(A_ScriptHwnd) applies only to the main window of current script

^Tab:: { ; cycle through main window views like browser tabs

winID := WinActive(".ahk - AutoHotkey v ahk_class AutoHotkey")
Text := WinGetText()

If RegExMatch(Text, "^Script lines most recently executed") {               ; Found - ListLines
    If A_ScriptHwnd == winID
        ListVars                ; ListVars - Variables and their contents   ; ^v
    Else Send "^v"
    ToolTipFn("2 ListVars - Variables and their contents", 5000, 150, -25)  ; 5s ; add # to avoid confusion
    }
Else If RegExMatch(Text, "^Local Variables|^Global Variables") {            ; Found - ListVars
    If A_ScriptHwnd == winID
        ListHotkeys             ; ListHotkeys - Hotkeys and their methods   ; ^h
    Else Send "^h"
    ToolTipFn("3 ListHotkeys - Hotkeys and their methods", 5000, 150, -25)  ; 5s
    }
Else If RegExMatch(Text, "^Type\s+Off\?\s+Level\s+Running\s+Name") {        ; Found - ListHotkeys
    If A_ScriptHwnd == winID
        KeyHistory              ; KeyHistory - Key history and script info  ; ^k
    Else Send "^k"
    ToolTipFn("4 KeyHistory - Key history and script info", 5000, 150, -25) ; 5s
    }
Else If RegExMatch(Text, "^Window: ") {                                     ; Found - KeyHistory
    If A_ScriptHwnd == winID
        ListLines               ; ListLines - Lines most recently executed  ; ^L
    Else Send "^l"
    ToolTipFn("1 ListLines - Lines most recently executed", 5000, 150, -25) ; 5s
    }
Else {
    ToolTipFn(A_ThisHotkey ":: InStr Text not found!", 2000)                ; 2s
    Exit
    }
}

^+Tab:: { ; cycle through main window views in reverse
winID := WinActive(".ahk - AutoHotkey v ahk_class AutoHotkey")
Text := WinGetText()

If RegExMatch(Text, "^Script lines most recently executed") {               ; Found - ListLines
    If A_ScriptHwnd == winID
        KeyHistory              ; KeyHistory - Key history and script info  ; ^k
    Else Send "^k"
    ToolTipFn("4 KeyHistory - Key history and script info", 5000, 150, -25) ; 5s
    }
Else If RegExMatch(Text, "^Local Variables|^Global Variables") {            ; Found - ListVars
    If A_ScriptHwnd == winID
        ListLines               ; ListLines - Lines most recently executed  ; ^L
    Else Send "^l"
    ToolTipFn("1 ListLines - Lines most recently executed", 5000, 150, -25) ; 5s
    }
Else If RegExMatch(Text, "^Type\s+Off\?\s+Level\s+Running\s+Name") {        ; Found - ListHotkeys
    If A_ScriptHwnd == winID
        ListVars                ; ListVars - Variables and their contents   ; ^v
    Else Send "^v"
    ToolTipFn("2 ListVars - Variables and their contents", 5000, 150, -25)  ; 5s
    }
Else If RegExMatch(Text, "^Window: ") {                                     ; Found - KeyHistory
    If A_ScriptHwnd == winID
        ListHotkeys             ; ListHotkeys - Hotkeys and their methods   ; ^h
    Else Send "^h"
    ToolTipFn("3 ListHotkeys - Hotkeys and their methods", 5000, 150, -25)  ; 5s
    }
Else {
    ToolTipFn(A_ThisHotkey ":: InStr Text not found!", 2000)                ; 2s
    Exit
    }
}

#HotIf

;------------------------------------------------------------------------------
;  = Calculator (classic)

#HotIf WinActive("ahk_class CalcFrame")

!1:: ; Standard Calculator
!2:: ; Scientific
!3:: ; Programmer
!4:: ; Statistics
{
; focus is on editing history and Alt + 3 is pressed
If ControlGetClassNN(ControlGetFocus("A")) == "Edit1" AND ThisHotkey = "!3"
    WrapTextFn(WrapText_Leading3 , WrapText_Trailing3) ; enclose in round brackets  ()
Else ; switch calculator type and go to basic view
    Send ThisHotkey "^{F4}"
}


; Toggle Date Calculation view
^d:: {
If checkCalcView() = 1
    Send "^{F4}"    ; Ctrl + F4 = Basic view
Else Send "^e"      ; Ctrl + E  = Date Calculation view
}

; Toggle Unit conversion view
^u:: {
If checkCalcView() = 2
    Send "^{F4}"    ; Ctrl + F4 = Basic view
Else Send "^u"      ; Ctrl + U  = Unit conversion view
}

^w::^F4     ; Go to Basic view

^Tab:: {    ; cycle through calculator views like browser tabs

CalcView := checkCalcView()

If CalcView = 0             ; Basic view
    Send "^u"               ; Ctrl + U  = Unit conversion view
Else If CalcView = 1        ; Unit conversion view
    Send "^e"               ; Ctrl + E  = Date Calculation view
Else If CalcView = 2        ; Date Calculation view
    MenuSelect , , "View", "Worksheets", "Mortgage"
Else If CalcView = 3        ; Worksheets > Mortgage
    MenuSelect , , "View", "Worksheets", "Vehicle lease"
Else If CalcView = 4        ; Worksheets > Vehicle lease
    MenuSelect , , "View", "Worksheets", "Fuel economy (mpg)"
Else If CalcView = 5        ; Worksheets > Fuel economy (mpg)
    MenuSelect , , "View", "Worksheets", "Fuel economy (L/100 km)"
Else                        ; Worksheets > Fuel economy (L/100 km)
    Send "^{F4}"            ; Ctrl + F4 = Basic view
}

^+Tab:: {

CalcView := checkCalcView()

If CalcView = 0             ; Basic view
    MenuSelect , , "View", "Worksheets", "Fuel economy (L/100 km)"
Else If CalcView = 1        ; Unit conversion view
    Send "^{F4}"            ; Ctrl + F4 = Basic view
Else If CalcView = 2        ; Date Calculation view
    Send "^u"               ; Ctrl + U  = Unit conversion view
Else If CalcView = 3        ; Worksheets > Mortgage
    Send "^e"               ; Ctrl + E  = Date Calculation view
Else If CalcView = 4        ; Worksheets > Vehicle lease
    MenuSelect , , "View", "Worksheets", "Mortgage"
Else If CalcView = 5        ; Worksheets > Fuel economy (mpg)
    MenuSelect , , "View", "Worksheets", "Vehicle lease"
Else                        ; Worksheets > Fuel economy (L/100 km)
    MenuSelect , , "View", "Worksheets", "Fuel economy (mpg)"
}

#HotIf
;  = Firefox

#HotIf WinActive("ahk_class MozillaWindowClass") ; main window ; excludes other dialogue boxes like "Save As" originating from ahk_exe firefox.exe

; Ctrl + Shift + F = close Find Bar
^+f::Send "^f{Esc}"

; Alt + H = open Homepage
!h::Send "^w^t"
; Background image loads correctly. Go backwards and Go forwards button history is lost, but not permanently. Use ^+t to restore most recent closed tab and check tab history OR use check recent history in Firefox Library.
; Send "^Labout:home{Enter}" ; alternative - Go backwards and Go forwards button history is preserved, but blank grey background may be seen instead of new tab background image

; Ctrl + Shift + O = open library / bookmark manager
^+o:: {
If WinActive(" — Mozilla Firefox") ; If not new tab, then open new one
    Send "^t"
Else Send "^l"  ; If new tab, focus address bar
Sleep 250       ; 250ms ; wait for focus - change as per your system performance
Send "{Raw}chrome://browser/content/places/places.xhtml")
Send "{Enter}"
}

; Ctrl + Shift + Q = Exit (Disable default Firefox shortcut)
^+q::Return

#HotIf


;------------------------------------------------------------------------------
;  = KeePass

#HotIf WinActive("ahk_exe KeePass.exe")

^n::^i ; not new database, new key entry

#HotIf

;------------------------------------------------------------------------------
;  = Dark Mode - Window Spy

#HotIf WinActive("Window Spy for AHKv2 ahk_class AutoHotkeyGUI")

^c:: {
A_Clipboard := StrReplace(CallClipboardVar(2), "`r`n", "`s") ; 2s, Exit ; copy contents to one line
}

#HotIf

;------------------------------------------------------------------------------
;  = Windows File Explorer

#HotIf WinActive("ahk_class CabinetWClass")

;    + Explorer main window

F1::F2 ; disable opening help in MS edge

;--------
;      * Unselect
; Unselect multiple files/folders
; Source: https://superuser.com/questions/78891/is-there-a-keyboard-shortcut-to-unselect-in-windows-explorer

; Ctrl + Shift + A = unselect by sending {F5} key ; same as Right Click > Refresh
^+a::F5
; alternative - WindowsRefreshOrRun()

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
    ; Sleep 500           ; alternative to WinWait ; see Note 2
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

If WinActive("Recycle Bin - File Explorer ahk_class CabinetWClass ahk_exe explorer.exe") {
    ; obtain drive letters, Breakdown per drive, total no of files, total size of files, list of files
    ; and display results in MsgBox
    RBinDisplay(RBinQuery())
    }
Else {
    ; check keyboard focus
    ClassNN := ControlGetClassNN(ControlGetFocus("A"))

    ; If keyboard focus = file list
    If ClassNN ~= "DirectUIHWND" ; RegEx match
        path := ValidPath()

    ; If keyboard focus = navigation pane
    Else If ClassNN == "SysTreeView321" {
        ToolTipFn("Focus is in Navigation Pane! GetFolderSize Aborted!", 2000)  ; 2s
        Send "{F6}{Home}"                                                   ; Return focus to file list
        CheckControlRegEx("DirectUIHWND")                                   ; ~= force focus File List
        Exit
        }

    Else path := [2, GetExplorerPath()] ; other ClassNN

    ; calculate folder size and display
    GetFolderSize(path[1], path[2])
    }
}

;--------
;      * Delete Empty Folder

^+d::DeleteEmptyFolder(CaptureFolderPath())

;--------
;      * Extract from folder & delete

^+e:: {

; extract folder path using clipboard
SourceFolder := CaptureFolderPath("`nSource Folder")            ; e.g., D:\Movies

; remove folder "Movies" and determine "D:\" is the destination
DestinationFolder := RegExReplace(SourceFolder, "\\[^\\]+$")    ; Result → D:\

If not DirExist(DestinationFolder) {
    MsgBox "Source: " SourceFolder "`nDestination: " DestinationFolder "`nDirectory does not exist!",, 262144 ; 262144 = Always-on-top
    Exit
    }

; move files
Try FileMove SourceFolder "\*.*", DestinationFolder
Catch as Err
    ErrorMsg := "FileMove`nSource: " SourceFolder "`nDestination: " DestinationFolder "`nErrors: " Err.Extra

; move folders
Loop Files, SourceFolder "\*.*", "D" {
    Try DirMove A_LoopFilePath, DestinationFolder "\" A_LoopFileName
    Catch {
        If IsSet(ErrorMsg)
            ErrorMsg .= "`nDirMove Error #" A_Index ": " A_LoopFilePath
        Else ErrorMsg := "DirMove Error #" A_Index ": " A_LoopFilePath
        }
    }

; show error or delete empty folder
If IsSet(ErrorMsg)
    MsgBox ErrorMsg,, 262144 ; 262144 = Always-on-top
Else DeleteEmptyFolder(SourceFolder)
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

;--------
;    + Locate desktop background

#HotIf WinActive("ahk_class WorkerW ahk_exe explorer.exe")

#w:: { ; win + w
path := WallpaperPath_v4()
result := MsgBox_Custom(path, "Current wallpaper location", 256 + 262144, 3, "In Explorer?", "Next Pic?", "OK") ; = Yes / No / Cancel ; 256 Button2 default ; 262144 = Always-on-top

; Actions
If result == "Yes" ; Button1 renamed to "In Explorer?"
    ; open path in windows file explorer and select file
    Run 'explorer.exe /n,/e,/select, ' '"' path '"'

Else If result == "No" { ; Button2 renamed to "Next Pic?"

    ; ask slideshow to go to next wallpaper
    nxtBackground()

    ; delete old wallpaper
    Try FileRecycle path
    Catch as err { ; If error, open photo in default app
        MsgBox A_ThisHotkey ":: Wallpaper deletion failed!`n" path "`n" Type(err) " - " err.What "`n" err.Message,, 262144 ; 262144 = Always-on-top
        Try Run path
        Catch as err ; If error, copy path to clipboard
            MsgBox A_ThisHotkey ":: Wallpaper deletion failed!`n" path "`n" Type(err) " - " err.What "`n" err.Message,, 262144 ; 262144 = Always-on-top
        }
    }
}

#HotIf

;------------------------------------------------------------------------------
; #HotIf Groups

;  = Capitalise the first letter of a sentence
; modified from https://www.autohotkey.com/board/topic/132938-auto-capitalize-first-letter-of-sentence/?p=719739

#HotIf WinActive("ahk_group CapitaliseFirstLetter_optIn")
; Below function is ENABLED in windows belonging to this group

; ~NumpadEnter:: ; disabled by default because of too many false positives
; ~Enter::       ; disabled - uncomment to use
~NumpadDot::
~.::
~!::
~?::
; ~？:: ; Unicode U+003F QUESTION MARK
; Triggers       ; add or disable one or more as needed
{
cfc1 := InputHook("L1 V C","{Space}{LShift}{RShift}{CapsLock}", "a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,x,y,z") ; captures 1st character, visible, case sensitive ; .a → .A
cfc1.Start
cfc1.Wait
If cfc1.EndReason == "Match" {
    If ThisHotkey == "~." OR ThisHotkey == "~NumpadDot"
    ; in case NumpadDot or . is the trigger, then don't capitalise coz typing the website address and file names is problematic ;  Example.exe → Example.exe (no change)
        Send "{BS}" cfc1.Input

    ; Else If ThisHotkey == "~NumpadEnter" OR ThisHotkey == "~Enter"
    ; in case NumpadEnter or Enter is the trigger, capitalise 1st character BUT don't add space, coz space is not necessary when creating a new paragraph
    ; commented out because trigger is disabled
    ;     Send "{BS}+" cfc1.Input

    Else { ; in case ! or ? is the trigger, then add a space and capitalise 1st character ; !a → ! A  and ?b → ? B
        Send "{BS} +" cfc1.Input
        ; SoundBeep 1500, 50 ; play a sound when successful - Frequency(a number between 37 and 32767), Duration in milliseconds
        ; SoundPlay "C:\Windows\Media\Windows Information Bar.wav" ; alternative to SoundBeep
        ; SoundPlay A_WinDir "\Media\Windows Balloon.wav"          ; alternative to SoundBeep
        }
    Exit ; stop capture if match found
    }
If cfc1.EndKey == "Space" { ; prevent cfc2 from firing for numbers or symbols. Example: 0.2ms is not changed to 0.2Ms
    cfc2 := InputHook("L1 V C","{Space}{LShift}{RShift}{CapsLock}", "a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,x,y,z") ; captures 2nd character, visible, case sensitive ; . a → . A
    cfc2.Start
    cfc2.Wait
    If cfc2.EndReason == "Match"
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
; Test whether method #2 works using Win + Shift + Wheel Up/Down keys (3-key combo)
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
#HotIf WinActive("ahk_group MediaKeysRestored")

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
; Hotstrings - Actions

;  = Date & Time

;    + Format Date / Time

:*x:d++::       Send FormatTime(, "yyyy.MM.dd")             ; sends 2021.02.31
:*x:date+::     Send FormatTime(, "dd.MM.yyyy")             ; sends 28.03.2020
:*x:time+::     Send FormatTime(, "h:mm tt")                ; sends 6:48 PM
:*x:datetime+:: Send FormatTime(, "dd.MM.yyyy h:mm tt")     ; sends 28.03.2020 6:46 PM

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
; Hotstrings - Text Replacement

;  = Find & Replace in Clipboard

;    + Find & Replace dot with space (StrReplace)

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
;  = Trim and change text

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

MyNotificationGui(mytext, myduration := 500, xAxis := 1550, yAxis := 985, timer := 1) { ; 500ms
Global MyNotification := Gui("+AlwaysOnTop -Caption +ToolWindow")   ; +ToolWindow avoids a taskbar button and an Alt-Tab menu item
MyNotification.SetFont("s9 w1000", "Arial")         ; font size 9, bold

; For dark mode
MyNotification.BackColor := "2C2C2E"                ; "2C2C2E" for dark mode background
MyNotification.AddText("cWhite w230 Left", mytext)  ; "cWhite" for dark mode text

; For light mode, comment out the above dark mode commands and uncomment below command -
; MyNotification.AddText("w230 Left", mytext) ; light mode text (background unchanged)

MyNotification.Show("x1650 y985 NoActivate")        ; NoActivate avoids deactivating the currently active window
WinMove xAxis, yAxis,,, MyNotification
If timer = 1
    SetTimer () => EndMyNotif(), Abs(myduration) * -1 ; 500ms ; new thread ; always negative number
If timer = 0 {
    Sleep Abs(myduration)                           ; always positive number
    EndMyNotif()
    }
}

;--------
;    + EndMyNotif

EndMyNotif() {
MyNotification.Destroy()
}

;------------------------------------------------------------------------------
;  = AHK Dark Mode functions

;    + ahkDarkMenu
/* primary source: https://stackoverflow.com/a/58547831/894589
   with modifications by
     * lexikos https://www.autohotkey.com/boards/viewtopic.php?p=426482#p426482
     * mcd https://www.autohotkey.com/boards/viewtopic.php?p=511756#p511756
*/

ahkDarkMenu() {
    static uxtheme := DllCall("GetModuleHandle", "str", "uxtheme", "ptr")
    static SetPreferredAppMode := DllCall("GetProcAddress", "ptr", uxtheme, "ptr", 135, "ptr")
    static FlushMenuThemes := DllCall("GetProcAddress", "ptr", uxtheme, "ptr", 136, "ptr")

    DllCall(SetPreferredAppMode, "int", 1) ; 0 = Default, 1 = AllowDark, 2 = ForceDark, 3 = ForceLight, 4=Max
    DllCall(FlushMenuThemes)
}

;------------------------------------------------------------------------------
;  = Toggle protected operating system (OS) files
; inspiration from https://www.autohotkey.com/board/topic/82603-toggle-hidden-files-system-files-and-file-extensions/?p=670182

;    + ToggleOS

ToggleOS(*) {
; alternative - Run ToggleSystemFiles.bat as administrator to toggle settings - https://superuser.com/a/1151851/391770
If Status = 0 { ; enable if disabled
    RegWrite "1", "REG_DWORD", "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSuperHidden"
    CheckRegWrite(Status)
    ToggleOSCheck
    SetTimer () => WindowsRefreshOrRun(), -100       ; 100ms ; new thread
    }
Else { ; disable if enabled
    RegWrite "0", "REG_DWORD", "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSuperHidden"
    CheckRegWrite(Status)
    ToggleOSCheck
    SetTimer () => WindowsRefreshOrRun(), -100       ; 100ms ; new thread
    }
}

;--------
;    + CheckRegWrite

CheckRegWrite(value) { ; check if RegWrite was success
Global Status := RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSuperHidden")
If value == Status {
    MsgBox "ToggleOS Failed",, 262144 ; 262144 = Always-on-top
    ; ToolTipFn("ToggleOS Failed", 1000) ; 1s, use tooltip and exit as an alternative to MsgBox
    Exit
    }
}

;--------
;    + ToggleOSCheck

ToggleOSCheck() { ; tray tick mark
If not IsSet(Status)
    Global Status := RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSuperHidden")
If Status = 0
    A_TrayMenu.UnCheck "&Toggle OS files"
Else A_TrayMenu.Check "&Toggle OS files"
}

;--------
;    + WindowsRefreshOrRun

WindowsRefreshOrRun() {
; Sleep 2000                                  ; 2s ; uncomment If necessary - wait for registry change to be enforced
If WinExist("ahk_class CabinetWClass") {    ; If Windows File Explorer window exists
    RefreshExplorer()
    If not WinActive()
        WinActivate
    }
Else { ; open new explorer window if one doesn't already exist ; comment out this section if not desired
    Run 'explorer.exe',,"Max"
    If WinWait("ahk_class CabinetWClass",, 10) ; 10s timeout
        If not WinActive()
            WinActivate
    }
}

;--------
;    + RefreshExplorer
; Source: https://www.autohotkey.com/boards/viewtopic.php?p=482766#p482766

RefreshExplorer() {
    Local Windows := ComObject("Shell.Application").Windows
    Windows.Item(ComValue(0x13, 8)).Refresh()   ; 0x13 = VT_UI4 32-bit unsigned int
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
UpdateWindows() {
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
If hwnd := WinExist("ahk_class CabinetWClass") {            ; if explorer is open
    WinActivate

    ; if paths are equal, no further action is needed
    If path = GetExplorerPath(hwnd)
        Return

    If WinWaitActive(,, 2) {                        ; 2s = Sleep 2000, but sends next command as soon as activated, instead of waiting for the full 2000ms period

        ; wait for cursor focus
        If FocusExplorerAddressBar() == "err0r" {
            Run 'explore "' path '"',,"Max"
            Exit
            }

        ; check to see if existing path is not equal to new path
        Else {
            Send "{Raw}" path
            WinWaitClose(,, 2)                      ; 2s - wait for drop down to disappear, then Send Enter ; WinWait commands used to prevent drop down display appearing after Enter - explorer bug
            Send "{Enter}{F6 2}"                    ; focus file list
            }
        }
    Else Run 'explore "' path '"',,"Max"
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
    If A_Index > 5 {                                        ; = Sleep 500ms ; wait until focus is on address bar
        ToolTipFn("Failed to focus address bar!", 2000)     ; 2s
        Return "err0r"
        }
    }
If WinWait("PopupHost ahk_class Microsoft.UI.Content.PopupWindowSiteBridge",, 2) ; 2s - wait for drop down
    Return "Success"
Else {
    ToolTipFn("Pop-up failed to appear!", 2000)     ; 2s
    Return "err0r"
    }
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
If Transparency == "Off"
    WinSetTransparent 255, "ahk_id " id
    ; Set transparency to 255 before using Off - might avoid window redrawing problems such as a black background. If the window still fails to be redrawn correctly, try WinRedraw, WinMove or WinHide + WinShow for a possible workaround.
WinSetTransparent Transparency, "ahk_id " id
ToolTipFn("Transparency: " Transparency) ; 500ms
}

;--------
;    + SetTransMenuFn
; modified from http://www.computoredge.com/AutoHotkey/Downloads/Always_on_Top.ahk

SetTransMenuFn() {
MouseGetPos ,, &WinID       ; identify window id
; WinID := WinExist("A")    ; alternative - but 'Active' window might not always be the intended target
Global WinID                ; so that SetTransByMenu can use it to set transparency
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
ToolTipFn("Transparency: " Trim(SubStr(item, 4)), 2000) ; 2s
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
Loop Parse A_Clipboard {                        ; Code Credit #2
    If StrLower(A_LoopField) == A_LoopField     ; * Code Credit #3
        inverted .= StrUpper(A_LoopField)       ; *
    Else inverted .= StrLower(A_LoopField)      ; *
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
Send "^v"       ; Paste text
If Len <= 20    ; and select text only if text ≤ 20 characters (change limit as needed)
    Send "+{Left " Len "}"
}

; Code Credit #1 - NeedleRegEx pattern modified from https://www.autohotkey.com/board/topic/24431-convert-text-uppercase-lowercase-capitalized-or-inverted/?p=158295
; Code Credit #2 - idea for loop modified from https://www.autohotkey.com/boards/viewtopic.php?p=58417#p58417
; Code Credit #3 - 3 lines of code with a comment "; *" were adapted from a (inaccurate) answer generated from a query to KudoAI's DuckDuckGPT user script - https://greasyfork.org/en/scripts/459849-duckduckgpt

;------------------------------------------------------------------------------
;  = Clipboard functions

;    + CallClipWait

CallClipWait(secs := 2, retrn := 0) {
ToolTipFn("Waiting for clipboard", secs * 1000)        ; 2s
If not ClipWait(secs) {
    ToolTipFn(A_ThisHotkey ":: Clip Failed", 2000)     ; 2s
    ; MyNotificationGui(A_ThisHotkey ":: Clip Failed", 2000) ; 2s ; alternative to tooltip
    Exit
    }

If retrn = 1
    Return A_Clipboard
}

;--------
;    + CallClipboard

CallClipboard(secs := 2, retrn := 0) {
ToolTipFn("Waiting for clipboard", secs * 1000)     ; 2s
Global clipSave := ClipboardAll()                   ; Global = Return clipSave
A_Clipboard := ""
Send "^c"
If not ClipWait(secs) {
    ToolTipFn(A_ThisHotkey ":: Clip Failed", 2000)  ; 2s
    A_Clipboard := clipSave
    If retrn = 0
        Exit
    Else Return "err0r"                             ; If retrn = 1
    }
}

;--------
;    + CallClipboardVar

CallClipboardVar(secs := 2, retrn := 0) {   ; copied text is sent to variable, clipboard is restored
If CallClipboard(secs, retrn) == "err0r"    ; ClipChanged is turned on
    Return "err0r"
Else {
    clipped := A_Clipboard
    A_Clipboard := clipSave
    Return clipped
    }
}

;------------------------------------------------------------------------------
;  = ToolTip function

;    + ToolTipFn

ToolTipFn(mytext, myduration := 500, xAxis?, yAxis?) { ; 500ms
If not IsSet(WhichToolTip)
    Static WhichToolTip := 1    ; 1
Else {
    ToolTip(,,, WhichToolTip)   ; turn Off previous ToolTip
    WhichToolTip++              ; add 1 to variable
}

; If WhichToolTip variable exceeds 20
If WhichToolTip > 20            ; inbuilt limit of 20
    WhichToolTip := 1           ; reset to 1

ToolTip mytext, xAxis?, yAxis?, WhichToolTip
SetTimer () => ToolTip(,,, WhichToolTip), Abs(myduration) * -1 ; 500ms ; new thread ; always negative number
}

;------------------------------------------------------------------------------
;  = Wrap Text In Quotes or Symbols
; Inspired by https://geekdrop.com/content/super-handy-autohotkey-ahk-script-to-change-the-case-of-text-in-line-or-wrap-text-in-quotes
; and https://www.autohotkey.com/board/topic/9805-easy-encloseenquote/?p=61995

;    + WrapTextMenuFn

WrapTextMenuFn() {
WrapTextMenu := Menu()
WrapTextMenu.Delete
WrapTextMenu.Add("&1   `'  Single Quotation `'"     , WrapTextMenuSelectionFn)
WrapTextMenu.Add("&2   `" Double Quotation `""      , WrapTextMenuSelectionFn)
WrapTextMenu.Add("&3   (  Round Brackets )"         , WrapTextMenuSelectionFn)
WrapTextMenu.Add("&4   [  Square Brackets ]"        , WrapTextMenuSelectionFn)
WrapTextMenu.Add("&5   {  Flower Brackets }"        , WrapTextMenuSelectionFn)
WrapTextMenu.Add("&6   ``  Accent/Backtick ``"      , WrapTextMenuSelectionFn)
WrapTextMenu.Add("&7  `% Percent Sign `%"           , WrapTextMenuSelectionFn)
WrapTextMenu.Add("&8   ‘  Single Comma Quotation ’" , WrapTextMenuSelectionFn)
WrapTextMenu.Add("&9   “ Double Comma Quotation ”"  , WrapTextMenuSelectionFn)
WrapTextMenu.Add("&0  Remove all"                   , WrapTextMenuSelectionFn)
WrapTextMenu.Show
}

;--------
;    + WrapTextMenuSelectionFn

WrapTextMenuSelectionFn(item, position, WrapTextMenu) {
If position = 1
    WrapTextFn(WrapText_Leading1 , WrapText_Trailing1)      ; enclose in single quotation '' - ' U+0027 : APOSTROPHE
Else If position = 2
    WrapTextFn(WrapText_Leading2 , WrapText_Trailing2)      ; enclose in double quotation "" - " U+0022 : QUOTATION MARK
Else If position = 3
    WrapTextFn(WrapText_Leading3 , WrapText_Trailing3)      ; enclose in round brackets  ()
Else If position = 4
    WrapTextFn(WrapText_Leading4 , WrapText_Trailing4)      ; enclose in square brackets []
Else If position = 5
    WrapTextFn(WrapText_Leading5 , WrapText_Trailing5)      ; enclose in flower brackets {}
Else If position = 6
    WrapTextFn(WrapText_Leading6 , WrapText_Trailing6)      ; enclose in accent/backtick ``
Else If position = 7
    WrapTextFn(WrapText_Leading7 , WrapText_Trailing7)      ; enclose in percent sign %%
Else If position = 8
    WrapTextFn(WrapText_Leading8 , WrapText_Trailing8)      ; enclose in ‘’ - ‘ U+2018 LEFT & ’ U+2019 RIGHT SINGLE QUOTATION MARK {single turned comma & comma quotation mark}
Else If position = 9
    WrapTextFn(WrapText_Leading9 , WrapText_Trailing9)      ; enclose in “” - “ U+201C LEFT & ” U+201D RIGHT DOUBLE QUOTATION MARK {double turned comma & comma quotation mark}
Else                                                        ; position = 10
    WrapTextFn( ""               , ""  )                    ; remove above quotes
}

;--------
;    + WrapTextFn

WrapTextFn(q, p, direct := 0) {

; enable wrapping text when WrapTextFn is directly triggered in MSPaintApp when inserting/editing text element (ClassNN: RICHEDIT50W1)
; If not, then just Send the triggered hotkey
If direct = 1 AND WinActive("ahk_class MSPaintApp") AND ControlGetClassNN(ControlGetFocus("A")) !== "RICHEDIT50W1" {
    Send A_ThisHotkey
    Exit
    }
; Else proceed with wrapping text

CallClipboard(2)                                        ; 2s, Exit
TextString := StrReplace(A_Clipboard, "`r")             ; remove \r for StrLen
TextStringInitial := TextString                         ; save initial string for later
TextString := RegExReplace(TextString,"^\s+|\s+$")      ; RegEx remove leading/trailing ; \s = [\r\n\t\f\v ]
Len := StrLen(TextString)                               ; string length

; remove existing wrap characters, predefined Global variables for leading '"([{`%‘“ and trailing ”’%`}])"' characters
; Example: "Hello" → Hello     and  '"([{`%‘“Hello”’%`}])"' → Hello
; However, "Hello' → "Hello' ---- no change because leading and trailing characters don't match
Loop {
    If A_Index = 10
        Break

    ; only If wrap characters are the leading (position 1) and trailing characters (last position = length of string)
    ; Example: '"([{`%‘“Hel'lo”’%`}])"' → Hel'lo
    If InStr(TextString, WrapText_Leading%A_Index%) = 1 AND InStr(TextString, WrapText_Trailing%A_Index%, , -1) = Len {
        TextString := SubStr(TextString, 2, Len - 2)    ; remove wrap characters If found
        Len := StrLen(TextString)                       ; determine string length again
        A_Index := 0                                    ; and reset Loop to check for multiple wrapping characters If any
        }
    }
/* ; Alternative to Loop command
; this works even when string contains mixed leading and trailing wrap characters such as "Hello' → Hello
TextString := RegExReplace(TextString,'^[\[`'\(\{%`"“‘]+|^``',, &ReplacementCount)     ;"; remove leading  ['({%"“‘`  ; customise as your needs in WrapTextMenuFn and WrapText Keys
If ReplacementCount > 0 ; don't remove trailing character if leading character is not removed
    TextString := RegExReplace(TextString,'[\]`'\)\}%`"”’]+$|``$')     ;"; remove trailing ]')}%"”’`  ; customise as your needs in WrapTextMenuFn and WrapText Keys
*/

; add new wrapping characters and determine length again
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

; Send "{Raw}" TextString   ; send the string with quotes
A_Clipboard := TextString   ; pasting from clipboard is faster than send raw, especially for long strings
Send "^v"                   ; paste
If Len <= 20                ; and select text string if ≤ 20 characters (change limit as needed)
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
ToolTipFn("Decoding successful!", 2000) ; 2s
Return Url
}

;--------
;    + UrlEncode

UrlEncode(str, sExcepts := "-_.", enc := "UTF-8") {
validateURL := "^http(s|)://[\w-@:%~#&=\.\+\?]{2,256}\.[a-z]{2,6}($|/[\w-@:%~#&=/\.\+\?]*$)" ; modified from https://www.makeuseof.com/regular-expressions-validate-url/
If not RegExMatch(str, validateURL) {
    Try UrlDecode(str) ; url is already encoded
    Catch {
        MsgBox "ERROR!`n`nString: " str "`nRegEx: " validateURL "`n`nNot a valid URL.",, 262144 ; 262144 = Always-on-top
        Exit
        }
    Else ToolTipFn("URL already encoded!", 2000) ; 2s
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
    ToolTipFn("Valid URL! Encoding successful!", 2000) ; 2s
    Return encoded
    }
}

;------------------------------------------------------------------------------
;  = Kill All Instances Of An App

;    + GetKillTitles

GetKillTitles(HWNDs, titleLimit) {
msgList := ""
Loop HWNDs.Length {
    t := RegExReplace(WinGetTitle(HWNDs[A_Index]), "\s+", A_Space)   ; remove multiple \s = [\r\n\t\f\v ]

    If t == ""
        tempTitle := Format("{:5}", A_Index) " = #No Title`n"
    Else tempTitle := Format("{:5}", A_Index) " = " t "`n"

    ; rudimentary word wrap ; for more sophisticated solution (AHK v1) - https://www.autohotkey.com/boards/viewtopic.php?t=59461
    If StrLen(tempTitle) > 50
        msgList .= SubStr(tempTitle, 1, 50) "…`n"
    Else msgList .= tempTitle

    If A_Index = titleLimit {
        msgList .= "    … and more. Too many to show here! Check Title List!"
        Break
        }
    }
Return msgList
}

;--------
;    + GetKillTitlesFileList

GetKillTitlesFileList(HWNDs) {
FileList := ""
Loop HWNDs.Length {
    t := RegExReplace(WinGetTitle(HWNDs[A_Index]), "\s+", A_Space)   ; remove multiple \s = [\r\n\t\f\v ]

    If t == ""
        FileList .= Format("{:5}", A_Index) " = #No Title`n"
    Else FileList .= Format("{:5}", A_Index) " = " t "`n"
    }
Return FileList
}

;------------------------------------------------------------------------------
;  = Print Screen functions

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
        ToolTipFn(A_ThisHotkey ":: Snipping Tool Pop-up timed out", 2000) ; 2s
        Exit
        }
    }
Else {
    ToolTipFn(A_ThisHotkey ":: Screen Snipping Overlay timed out", 2000) ; 2s
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
        Sleep 500                                           ; 500ms
        Send "^v"                                           ; paste image from clipboard
        ; ControlSend "^v",, "ahk_class MSPaintApp"           ; alternative to Send
        }
 */
    }
Else If A_PriorKey == "Escape"
    ToolTipFn(A_ThisHotkey ":: Screen Snipping aborted - Esc", 2000) ; 2s
Else ToolTipFn(A_ThisHotkey ":: Screen Snipping aborted - 15s timeout", 2000) ; 2s
}

/* For older versions of Snipping tool, below code may work

SnipFromMenu(ItemName, ItemPos, MyMenu) {
Send "{PrintScreen}"

; wait for screenshot tool to activate
If WinWaitActive("Screen Snipping ahk_class Windows.UI.Core.CoreWindow",, 3) { ; 3s timeout
    Sleep 500
    Send "{Tab " ItemPos "}{Enter}"
    }
Else {
    ToolTipFn(A_ThisHotkey ":: Screen Snipping timed out", 2000) ; 2s
    Exit
    }

; wait for screenshot to be taken ; abort further action if timeout or Esc key is pressed
If WinWaitClose(,, 15) and (A_PriorKey != "Escape") { ; 15s
    If WinExist("ahk_class MSPaintApp")
        WinActivate
    Else Run "mspaint.exe",,"Max"

    If WinWait("ahk_class MSPaintApp",, 3) ; 3s
        ; modified from https://www.autohotkey.com/boards/viewtopic.php?p=360163#p360163
        Send "^v"                                           ; paste image from clipboard
        ; ControlSend "^v",, "ahk_class MSPaintApp"         ; alternative to Send
    }
Else ToolTipFn(A_ThisHotkey ":: Screen Snipping aborted - 15s timeout / Esc", 2000) ; 2s
}

*/

;--------
;    + PrintScreenFn

PrintScreenFn(*) {
; save screenshot number in variable 'serial'
serial := RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer", "ScreenshotIndex", "1")

; send Windows + Print Screen
Send "#{PrintScreen}"

; use SetTimer to wait for file creation 100ms and execute action in another thread
; pass serial as number for path of screenshot file in `ScreenshotFileOp` function
SetTimer () => ScreenshotFileOp("C:\Users\" A_UserName "\Pictures\Screenshots\Screenshot (" serial ").png"), -100   ; 100ms ; new thread
}

;--------
;    + ScreenshotFileOp

ScreenshotFileOp(MyPath) {
While not FileExist(MyPath) {   ; if file does not exist, wait using Sleep and repeat check
    Sleep 500                   ; wait for 500ms
    If A_Index > 6 {            ; max 6 attempts, total wait 3000ms = 3s
        ; If failed ×6, open screenshot folder in explorer
        OpenFolder("C:\Users\" A_UserName "\Pictures\Screenshots\")
        MsgBox MyPath,, 262144      ; 262144 = Always-on-top ; show calculated path in pop-up to compare with actual file path
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
;  = Check Window Focus/Control

;    + CheckControlRegEx

CheckControlRegEx(ClassNN := "DirectUIHWND") {
; default "DirectUIHWND2" = File List focus in File Explorer ; "DirectUIHWND3" = File List focus in file explorer changed by ExplorerPatcher.exe
If not ControlGetClassNN(ControlGetFocus("A")) ~= ClassNN {
    WinGetClientPos , , &OutWidth, &OutHeight
    MouseMove OutWidth * 0.75, OutHeight * 0.75 ; move mouse to bottom right quadrant (usually empty)
    Send "{Click}"
    }
}


;------------------------------------------------------------------------------
;  = Windows File Explorer Fn

;    + GetFolderSize

GetFolderSize(pathType, pathContent) {
; variables
SizeB := 0
errorDetails := ""
folderName := ""
fileName := ""

If pathType = 1 {   ; multiple lines ; InStr(pathContent, "`n")
    Loop Parse pathContent, "`n", "`r" {
        If DirExist(A_LoopField) {          ; is folder
            Loop Files A_LoopField "\*.*", "R"
                SizeB += A_LoopFileSize
            folderName .= A_LoopField "`n"
            }
        Else If FileExist(A_LoopField) {    ; is file
            SizeB += FileGetSize(A_LoopField, "B")
            fileName .= A_LoopField "`n"
            }
        Else errorDetails .= "#" A_Index ": " A_LoopField " → Not a folder or file`n"
        }

    ; display folders first in pathContent
    If folderName == ""
        pathContent := Trim(Sort(fileName, "N"), "`n")
    Else If fileName == ""
        pathContent := Trim(Sort(folderName, "N"), "`n")
    Else pathContent := Trim(Sort(folderName, "N"), "`n") "`n" Trim(Sort(fileName, "N"), "`n")

    ; limit string to 1000 characters
    If StrLen(pathContent) > 1500 {
        RegExMatch(pathContent, "[\s\S]{1,1500}`n", &output)
        pathContent := output.0 "… and more!"
        }
    }
Else If pathType = 2 { ;          ; single line - folder ; DirExist(pathContent)
    Loop Files pathContent "\*.*", "R"
        SizeB += A_LoopFileSize
    }
Else SizeB += FileGetSize(pathContent, "B")    ; single line - file ; FileExist(pathContent)

; check size and display tooltip
If SizeB = 0 and errorDetails != ""   ; size zero and error present
    ToolTipFn(errorDetails, 5000)          ; 5s ; show error
Else If SizeB = 0                     ; size zero and no errors
    ToolTipFn(pathContent "`nEmpty Folder/File", 3000) ; 3s
Else ToolTipFn(pathContent "`nSize: " SizeFn(SizeB) "`n`n" errorDetails, 3000) ; 3s
}

;--------
;    + RBinQuery
; Modified from https://www.autohotkey.com/boards/viewtopic.php?p=125495#p125495

RBinQuery(userDriveInput := "") {
Global err0r := "" ; store errors

RBinArr         := [] ; array to store return
driveLetters    := checkDrivesFn(userDriveInput)

If userDriveInput == ""
    userDriveInfo := "User Query - Blank"
Else userDriveInfo := 'User Query - "' userDriveInput '"'

If userDriveInput == "" and err0r == ""
    ErrorReturn := ""
Else If err0r == ""
    ErrorReturn := "No errors in user query!`n"
Else ErrorReturn := "Errors in user query:`n" err0r

If driveLetters == "" {
    MsgBox("No matches to drive letters found!`n`n"
         . userDriveInfo "`n"
         . ErrorReturn,, 262144) ; 262144 = Always-on-top
    Exit
    }

SID := userSID()

; list of folders/files visible in windows file explorer
visibleArr := RBinVisible(driveLetters) ; returns [visibleSummary, visibleRBFCount, visibleTotalSizeB]

; list of folders/files *NOT* visible in windows file explorer
hiddenArr := RBinHidden(driveLetters, SID, visibleArr[1]) ; returns [RBSummary, TotalRBFCount, TotalSizeB, hiddenFileList]

RBinArr.InsertAt(1, driveLetters)
RBinArr.InsertAt(2, visibleArr[1])                              ; visibleSummary
RBinArr.InsertAt(3, visibleArr[2])                              ; visibleRBFCount   (used for hiddenRBFCount)
RBinArr.InsertAt(4, SizeFn(visibleArr[3]))                      ; visibleTotalSizeB (used for hiddenTotalSizeB)

RBinArr.InsertAt(5, hiddenArr[1])                               ; RBSummary
RBinArr.InsertAt(6, hiddenArr[2])                               ; TotalRBFCount
RBinArr.InsertAt(7, SizeFn(hiddenArr[3]))                       ; TotalSizeB

RBinArr.InsertAt(8, hiddenArr[2] - visibleArr[2])               ; hiddenRBFCount
RBinArr.InsertAt(9, SizeFn(hiddenArr[3] - visibleArr[3]))       ; hiddenTotalSizeB
RBinArr.InsertAt(10, hiddenArr[4])                              ; hiddenFileList

RBinArr.InsertAt(11, userDriveInfo)                             ; userDriveInfo
RBinArr.InsertAt(12, ErrorReturn)                               ; ErrorReturn

Return RBinArr
}

;--------
;    + RBinVisible
; Source of original power shell code: https://stackoverflow.com/a/27769278/5282788
; Source of additional recycle bin item properties power shell code - https://jdhitsolutions.com/blog/powershell/7024/managing-the-recycle-bin-with-powershell/
; For ONLY this function - the power shell code was converted to AHK v2 by DuckDuckGo AI (Model: Chat GPT-4o mini @ https://duckduckgo.com/) on 2024.11.09
; The output was *very* FLAWED. It was corrected, further modified and tested by me.

RBinVisible(driveLetters) {

visibleFiles        := []  ; Create an empty array to hold the file information
visibleTotalSizeB   := 0   ; Create an empty variable to calculate total file size

; Use COM to access the Shell.Application
recycleBin := ComObject("shell.application").NameSpace(10) ; 10 is the identifier for the Recycle Bin

; Loop through each item in the Recycle Bin
For item in recycleBin.Items {

    If not RegExMatch(item.Path, "[" driveLetters "]+\:\\")
        Continue ; skip

    ; Add the information to the array
    visibleFiles.Push({
        Location: item.Path,
        OriginalLocation: item.ExtendedProperty("{9B174B33-40FF-11D2-A27E-00C04FC30871} 2"),
        Name: item.Name,
        FileSize: item.Size
        })

    visibleTotalSizeB += item.Size
    }

visibleRBFCount := visibleFiles.Length
visibleSummary := "; List of folders/files visible in File Explorer Recycle Bin -`n`n"
          . "Number of visible items     : " visibleRBFCount "`n"
          . "Total size of visible items : " SizeFn(visibleTotalSizeB) "`n`n"

For item in visibleFiles {
    visibleSummary .= "#" Format("{:03d}", A_Index) " Recycle Bin Location   : " item.Location
                   .  "`n     Original Location      : " item.OriginalLocation  "\" item.Name
                   .  "`n     File Size              : " SizeFn(item.FileSize) "`n`n"
    }

Return [visibleSummary, visibleRBFCount, visibleTotalSizeB]
}

;--------
;    + RBinHidden

RBinHidden(driveLetters, SID, visibleSummary) {

TotalRBFCount   := 0    ; total number of recycled files
TotalSizeB      := 0    ; total size in bytes
RBSummary       := ""

Loop StrLen(driveLetters) {
    RBFCount    := 0        ; Recycle Bin File count
    SizeB       := 0        ; size in bytes
    DriveL      := SubStr(driveLetters, A_Index, 1)
    Loop Files DriveL ":\$Recycle.Bin\" SID "\*.*", "R" {
        RBFCount++
        SizeB += A_LoopFileSize
        If not InStr(visibleSummary, A_LoopFileFullPath) ; exclude file name if already in visibleSummary
            FileNames .= Format("{:-80}", A_LoopFileFullPath) "(" SizeFn(A_LoopFileSize) ")`n"
        }
    RBSummary .= "Drive: " DriveL "`n"
              .  "No. of files: " RBFCount "`n"
              .  "Size: " SizeFn(SizeB) "`n`n"
    TotalRBFCount += RBFCount
    TotalSizeB += SizeB
    }

hiddenFileList := Trim(Sort(FileNames, "N"), "`n")
Return [RBSummary, TotalRBFCount, TotalSizeB, hiddenFileList]
}

;--------
;    + userSID

userSID() {
    If IsSet(SID)
        Return SID
    ; Else ; proceed

    Loop Reg "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList", "R KV" {
        If (A_LoopRegName == "ProfileImagePath") {
            If InStr(RegRead(), A_UserName) {
                Static SID := StrReplace(A_LoopRegKey, "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\")
                Return SID
                }
            }
        }
    Else {
        ToolTipFn(A_ThisHotkey ":: RBinQuery - userSID() error", 2000) ; 2s
        Exit
        }
}

;--------
;    + checkDrivesFn

checkDrivesFn(userDriveInput := "") {
Global err0r
driveList := DriveGetList()

If userDriveInput == ""        ; If drive letters are not specified or not valid
    Return driveList           ; get list of drives - e.g., ACDEZ
Else {
    checkedDrives := ""
    Loop Parse userDriveInput {
        If A_LoopField ~= "[" driveList "]"
            checkedDrives .= A_LoopField "`n"
        Else err0r .= ' * Bad Match - "' A_LoopField '" → Available drives: ' driveList "`n"
        }
    Return StrReplace(Sort(checkedDrives, "U"), "`n") ; remove U duplicates and `n line feed
    }
}

;--------
;    + RBinDisplay

RBinDisplay(RBinArr) {

msgTitle := "RBinQuery: " RBinArr[1]                                                  ; driveLetters
msgText := RBinArr[11] "`n"                                                           ; userDriveInfo
        .  RBinArr[12] "`n"                                                           ; ErrorReturn
        .  "# Below summary includes all files (visible + hidden)`n`n"
        .  RBinArr[5] "Total No. of files: " RBinArr[6] "`nSize: " RBinArr[7]         ; RBSummary, totals

mbcResult := MsgBox_Custom(msgText, msgTitle, 256 + 262144, 2, "File List", "OK")
; = Yes / No ; 256 Default2 262144 = Always-on-top

If mbcResult == "Yes" { ; renamed button - File List
    rbFile := A_Temp "\List of files in Recycle bin.txt"
    Results := RBinGenerateTxt(msgTitle, msgText, RBinArr)
    FileCreate_Or_Append(rbFile, Results)
    Run '"' rbFile '"',,"Max" ;"
    }
}

;--------
;    + RBinGenerateTxt

RBinGenerateTxt(msgTitle, msgText, RBinArr) {
Results := " ;" continuation
(
; * Table of contents *
; Summary
; List of folders/files visible in File Explorer Recycle Bin -
; List of hidden folders/files in Recycle Bin -

;------------------------------------------------------------------------------
; Summary


)" . msgTitle "`n`n" ;" 1 driveLetters
   . msgText  "`n`n" ; 5 RBSummary, 6 TotalRBFCount, 7 TotalSizeB
   . ";------------------------------------------------------------------------------`n"
   . RBinArr[2]      ; visibleSummary
   . ";------------------------------------------------------------------------------`n"
   . "; List of hidden folders/files in Recycle Bin -`n`n"
   . "Number of hidden items     : " RBinArr[8] "`n"        ; hiddenRBFCount
   . "Total size of hidden items : " RBinArr[9] "`n`n"      ; hiddenTotalSizeB
   . RBinArr[10]                                            ; hiddenFileList

return Results
}

;--------
;    + SizeFn

SizeFn(SizeB) {
If SizeB >= 1024 {
    readableSizeKB := SizeB / 1024
    If readableSizeKB >= 1024 {
        readableSizeMB := readableSizeKB / 1024
        If readableSizeMB >= 1024
            Return Round(readableSizeMB / 1024, 2) " GB"
        Else Return Round(readableSizeMB, 2) " MB"
        }
    Else Return Round(readableSizeKB, 2) " KB"
    }
Else Return SizeB " bytes"
}

;--------
;    + ValidPath

ValidPath() {
clipped := CallClipboardVar(2, 1) ; 2s, Return
If clipped == "err0r" {
    ToolTipFn(A_ThisHotkey ":: Error - Folder/file path copy failed!", 2000) ; 2s
    Exit
    }
If InStr(clipped, "`n") ; If multiple folder/files or single folder or file
    Return [1, clipped]
Else If DirExist(clipped)
    Return [2, clipped]
Else If FileExist(clipped)
    Return [3, clipped]
Else {
    MsgBox clipped "`nFolder/File does not exist!",, 262144 ; 262144 = Always-on-top
    Exit
    }
}

;--------
;    + DeleteEmptyFolder

DeleteEmptyFolder(WhichFolder) {
FolderSizeB := 0
contents := ""
Loop Files WhichFolder "\*.*", "R" {
    FolderSizeB += A_LoopFileSize
    contents .= A_LoopFileFullPath "`n"
    }
If contents == ""
    contents := "Nil"
If FolderSizeB = 0 {            ; If zero size
    FileRecycle WhichFolder     ; delete empty folder, recursive
    ToolTipFn(WhichFolder "`nEmpty folder deleted!", 1000) ; 1s
    }
Else ToolTipFn(WhichFolder . SizeFn(FolderSizeB) "`n`nContents:`n" contents, 5000) ; 5s
}

;--------
;    + CaptureFolderPath

CaptureFolderPath(errorTxt := "") {
clipped := CallClipboardVar(2, 1) ; 2s, Return
If clipped == "err0r" {
    ToolTipFn(A_ThisHotkey ":: Error - Folder path copy failed!" errorTxt, 2000) ; 2s
    Exit
    }
If not DirExist(clipped) {
    MsgBox clipped "`nNot a folder!" errorTxt,, 262144 ; 262144 = Always-on-top
    Exit
    }
Else Return clipped
}

;--------
;    + GetExplorerPath
; source: https://old.reddit.com/r/AutoHotkey/comments/10fmk4h/get_path_of_active_explorer_tab/kuplyts/

GetExplorerPath(hwnd := WinExist("A")) {
Try activeTab := ControlGetHwnd("ShellTabWindowClass1", hwnd)
Catch as err {
    ToolTipFn(A_ThisHotkey ":: Failed to get GetExplorerPath(hwnd) -`n" err.Message, 2000) ; 2s
    return false
    }
For w in ComObject("Shell.Application").Windows {
    If (w.hwnd == hwnd) {
        Static IID_IShellBrowser := "{000214E2-0000-0000-C000-000000000046}"
        shellBrowser := ComObjQuery(w, IID_IShellBrowser, IID_IShellBrowser)
        ComCall(3, shellBrowser, "uint*", &thisTab := 0)
        If thisTab == activeTab
            Return w.Document.Folder.Self.Path
        }
    }
ToolTipFn(A_ThisHotkey ":: Failed to get GetExplorerPath(w.hwnd) match!", 2000) ; 2s
Exit
}

;--------
;    + WallpaperPath_v4
; for more iterations of this function, visit this standalone script - https://github.com/xypha/AHK-v2-scripts/blob/main/standalone/WallpaperPath.ahk

WallpaperPath_v4(key := A_ThisHotkey) {

; Read the binary value from the registry
regBinary := RegRead("HKEY_CURRENT_USER\Control Panel\Desktop", "TranscodedImageCache") ; 1600 characters
; alternative: regBinary := RegRead("HKEY_CURRENT_USER\Control Panel\Desktop", "TranscodedImageCache_000")

; remove first 48 characters and create array
regArr := StrSplit(SubStr(regBinary, 49))

; variable that stores path
ConvString := ""

Loop regArr.Length {
    If Mod(A_Index, 4) != 0 ; generate `char` every 4th Loop, when mod() = 0 (division remainder)
        Continue            ; skip rest of the Loop if A_Index is 1 2 3 (not 4) 5 6 7 (not 8) 9…
    Else                    ; in Loop 4, 8… reposition characters as Chr("0x" [3] [4] [1] [2]), Chr("0x" [7] [8] [5] [6])… and so on
        char := Chr("0x" regArr[A_Index - 1] regArr[A_Index] regArr[A_Index - 3] regArr[A_Index - 2])

    If char = ""            ; char := Chr("0x" 0 0 0 0) consecutive zeroes results in empty string
        Break               ; End Loop
    Else ConvString .= char ; add generated string to path
    }
If FileExist(ConvString)
    Return ConvString           ; Return path to desktop wallpaper
Else {
    MsgBox key ":: WallpaperPath_v4() is not valid!`nConverted String: " ConvString "`n`nTranscodedImageCache:`n" regBinary,, 262144 ; 262144 = Always-on-top
    Exit
    }
}

;--------
;    + nxtBackground
; modified from https://www.autohotkey.com/boards/viewtopic.php?p=581885&sid=c79061ab577091808edda54ee99ed52e#p581885

nxtBackground() {
Try {
    pDesktopWallpaper := ComObject("{C2CF3110-460E-4fc1-B9D0-8A1C0C9CC4BD}", "{B92B56A9-8B55-4E14-9A89-0199BBB6F93B}")
    ComCall(16, pDesktopWallpaper, "Ptr", 0, "UInt", 0)
    ; ObjRelease(pDesktopWallpaper) ; causes error: ValueError - Parameter #1 of ObjRelease is invalid.
    pDesktopWallpaper := ""
    }
Catch as err
    ToolTipFn(A_ThisHotkey ":: nxtBackground failed!`n" Type(err) " - " err.What "`n" err.Message, 2000) ; 2s
}

;------------------------------------------------------------------------------
;  = Check if file exists and create/append

;    + FileCreate_Or_Append
; modified from AutoHotkey.chm::/docs/lib/FileOpen.htm#ExWriteRead

FileCreate_Or_Append(FilePath, FileContent) {
If FileExist(FilePath) {                                ; If file exists
        Try FileObj := FileOpen(FilePath, "w", "UTF-8") ; overwrite file contents
        Catch as Err {
            MsgBox FilePath "`nCan't open file for writing.`n`n" Type(Err) ": " Err.Message
            Return
            }
        FileObj.Write(FileContent)
        FileObj.Close()
        }
    Else FileAppend FileContent, FilePath, "`n UTF-8" ; create new file
}

;------------------------------------------------------------------------------
;  = Change MsgBox button text

;    + MsgBox_Custom
; Modified from AutoHotkey.chm::/docs/scripts/index.htm#MsgBoxButtonNames

MsgBox_Custom(MsgText := "Text", MsgTitle := A_ScriptName, msgOption := 262144, nButtons := 1, bName1 := "Button_Name", bName2?, bName3?) { ; 262144 = Always-on-top ; nButtons = 1/2/3 button options

If nButtons = 1 {
    msgOption += 0 ; OK button
} Else If nButtons = 2 and isSet(bName2) {
    msgOption += 4 ; YesNo button
} Else If nButtons = 3 and isSet(bName2) and isSet(bName3) {
    msgOption += 3 ; YesNoCancel button
} Else {
    MsgBox A_ThisHotkey "::" A_ThisFunc "`nError! Check options and button number and names",, 262144 ; 262144 = Always-on-top
    Exit
    }

RandomMsg := Random(1, 194161877)
SetTimer () => MsgBox_ChangeButtonText(RandomMsg, MsgTitle, nButtons, bName1, bName2?, bName3?), -100 ; 100ms ; new thread
Return MsgBox(MsgText, RandomMsg, msgOption)
}

;--------
;    + MsgBox_ChangeButtonText

MsgBox_ChangeButtonText(RandomMsg := 0, MsgTitle := A_ScriptName, nButtons := 1, bName1 := "OK", bName2?, bName3?) {
err0r := ""
If RandomMsg = 0 {
    If WinWait(MsgTitle " ahk_class #32770",, 2)        ; 2s
        SetButtonNames()
    Else err0r .= ":: MsgBox_ChangeButtonText failed! MsgTitle - " MsgTitle ": Timed out!"
    }
Else If WinWait(RandomMsg " ahk_class #32770",, 2) {    ; 2s
    WinSetTitle MsgTitle
    SetButtonNames()
    }
Else err0r .= ":: MsgBox_ChangeButtonText failed! RandomMsg - " RandomMsg ": Timed out!"

If err0r != ""
    ToolTipFn(A_ThisHotkey "`n" err0r, 2000) ; 2s
Else Return

    SetButtonNames() {
    ControlSetText bName1, "Button1"

    If isSet(bName2)
        ControlSetText bName2, "Button2"

    If isSet(bName3)
        ControlSetText bName3, "Button3"
    }
}

;------------------------------------------------------------------------------
;  = Calculator view (classic)

;    + checkCalcView

checkCalcView() {

DetectHiddenText False
WinText := WinGetText("ahk_class CalcFrame")
DetectHiddenText True

If InStr(WinText, "Select the type of unit you want to convert")
    CalcView := 1       ; Unit conversion view
Else If InStr(WinText, "Select the date calculation you want")
    CalcView := 2       ; Date Calculation view
Else If InStr(WinText, "Down payment") OR InStr(WinText, "Purchase price")
    CalcView := 3       ; Worksheets > Mortgage
Else If InStr(WinText, "Lease value") OR InStr(WinText, "Residual value")
    CalcView := 4       ; Worksheets > Vehicle lease
Else If InStr(WinText, "Fuel ") AND not RegExMatch(WinText, "liters|kilometers")
    CalcView := 5       ; Worksheets > Fuel economy (mpg)
Else If InStr(WinText, "Fuel ") AND not RegExMatch(WinText, "gallons|miles")
    CalcView := 6       ; Worksheets > Fuel economy (L/100 km)
Else CalcView := 0      ; Basic view

Return CalcView
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

; CLSID List (Windows Class Identifiers)                           ; Find more in AutoHotkey.chm::/docs/misc/CLSID-List.htm
::{21EC2020-3AEA-1069-A2DD-08002B30309D}                           ; Control Panel (view: small icons) ; alt
::{26EE0668-A00A-44D7-9371-BEB064C98683}                           ; Control Panel (view: category) ; alternate
::{20D04FE0-3AEA-1069-A2D8-08002B30309D}                           ; This PC

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
; * Test

:*:test++:: {
MsgBox ThisHotkey ":: Not assigned!",, 262144 ; 262144 = Always-on-top
}

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

v2.12 - 2024.10.31
 - remove `%A_ScriptDir%` from #include commands. It is already built in
 * rename `EncText` to `WrapTextFn`, `WrapTextFromMenu` to `WrapTextMenuSelectionFn`
 * improve all `WrapText` functions by adding Global variables for easy customisation
 * fix `WrapTextFn` by adding `Loop` command, because incorrect removal of leading/trailing characters under the assumption of mixed wrap characters. Note: keep old RegExReplace command in case some users want this behaviour.
 * fix `CallClipboard` incorrect ToolTip after successful copy by moving it prior to sending `^c`
 * fix `ToolTipFn` - failed to assign 1 and 20 to `WhichToolTip` variable, and hence failure to turn Off some ToolTips
 * rearrange/rename some function headings and update TOC
 * improve comments and small changes

v2.13 - 2024.11.08
 ★ add new shortcuts for `AHK Main Window` - cycling views through `Tab and `+Tab` (like browser tabs)
 ★ add shortcut for copying text into a single line when using `Dark Mode - Window Spy`
 ★ add shortcut for `Delete Empty Folder` in Windows File Explorer -- `^+d`
 ★ add shortcut for `Extract from folder & delete` -- `^+e` -- this extracts all items(folders & files) from selected folder to parent folder and deletes the file after emptying it
 ★ add `ahkDarkMenu()` to enable dark mode for `Menu()`
 ★ add several functions that allow querying recycle bin for size and folders/files; display results and generate reports -- see shortcut `^+s` or function `RBinDisplay()`
 ★ add new function `FileCreate_Or_Append` to open existing file If it exists to given path or open a new file and write contents of provided input
 ★ add new function `MsgBox_Custom` to create custom MsgBox and rename 1/2/3 built-in buttons as required
 * rebuild `GetKillTitles` function - add variable to limit size of MsgBox; add file list button If limit is exceeded and function `GetKillTitlesFileList` for it; improve displayed messages
 * rename function and variables within - `FolderSizeFn` to `SizeFn` ; `FolderSizeB` to `SizeB` and so on
 - remove "Size: " text additions in `SizeFn`. These can be added by calling function after returning results
 * improve `Sort` commands by adding numerical sorting to match file sorting in windows file explorer as applicable
 - remove negative numbers for Sleep and SetTimer commands from user-defined functions by using `Abs()` command and pre-assigned sign
 * improve MsgBox commands by changing options from string to numbers (math)
 * change shortcut to open Firefox Homepage from `^+h` to `!h`, and add explanation in comments
 * change "=" to "==" (again) wherever applicable to enable case sensitivity
 * improve `WindowsRefreshOrRun()` - comment out unnecessary `Sleep` command; rearrange/modify remaining commands
 * replace `!` in If commands with `not` to improve readability
 - remove unnecessary variable `errorTxt` from `ValidPath()`
 * rearrange/rename/update headings in TOC
 * improve comments and other small changes

v2.14 - 2024.11.11
 * change the nature of `Capitalise first letter` section from opt-out to opt-in. Renamed group name from `CapitaliseFirstLetter` to `CapitaliseFirstLetter_optIn` and changed the example.
 * change behaviour of capitalisation when dot/numdot is followed immediately by alphabet - stop inserting Space to allow for trigger-free typing of file extensions
 * renamed group `MediaKeys` to `MediaKeysRestored` for clarity and added comment on UIAccess
 * for users running explorerPatcher.exe, to prevent explorer shortcuts from breaking, change case sensitive matching to RegEx matching for ClassNN, and create new function `CheckControlRegEx`
 * improve `FocusExplorerAddressBar()` to show tooltip on failure
 * improve comments and other small changes

v2.15 - 2024.11.20
 ★ add `Calculator (classic)` section to #HotIf Apps; along with `checkCalcView` function
 ★ add `KeePass` section to #HotIf Apps
 ★ add "Locate desktop background" shortcut Win + W to `Windows File Explorer` section of #HotIf Apps; along with associated `WallpaperPath_v4` and `nxtBackground` functions
 * replace some mediocre code in "Show folder/file size in ToolTip" section and `OpenFolder` function with `GetExplorerPath` function for faster performance
 * improve `AHK Dark Mode` section by adding registry check before running `ahkDarkMenu`
 * rename "ahk_group NoWrapText" to "ahk_group WrapText_disabled" to maintain consistency
 * commented out `MSPaintApp` from "ahk_group WrapText_disabled" and change `WrapTextFn` to enable wrapping text when inserting/editing text element (ClassNN: RICHEDIT50W1) in classic paint
 * replace `MSPaintApp` with new example `ahk_class CalcFrame` for "ahk_group WrapText_disabled"
 * replace `/` (forward slash) with `.` (fullstop) in `datetime+` hotstring, inline with similar hotstrings in "Format Date / Time" section
 * replace `HKCU` with `HKEY_CURRENT_USER` and `HKLM` with `HKEY_LOCAL_MACHINE` wherever applicable
 * improve `CallClipWait` function to return clipboard if successful
 * improve "Clipboard functions" by adding default value (2 seconds) to `secs` variable
 * fix `SnipFromMenu` function by replacing non-functional `PostMessage` command with `Send`
 ! enabled "UTF-8" encoding for `FileAppend` command, instead of the previous "CP0" (system default ANSI) encoding. Might cause error messages if you are using the `FileCreate_Or_Append` function outside this script
 * rearrange/rename/update headings in TOC
 * improve comments and other small changes
*/
