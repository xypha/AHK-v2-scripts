; https://github.com/xypha/AHK-v2-scripts/edit/main/Showcase.ahk
; Last updated 2025.05.25

; Visit AutoHotkey (AHK) version 2 (v2) documentation to search for commands and in-built functions used in this script @
; https://www.autohotkey.com/docs/v2/

; comments begin with semi-colon ";" at start of line or space or after code in middle of line
; comments can also be enclosed by `/* */`, like this - /* comment text */
; and these two methods can be combined too :)

;   /* AHK v2 Showcase - CONTENTS */
; Settings
; Auto-execute
;  = Set Global variables
;    + MenuShortcuts
;    + MenuIcons_path
;    + Send_LenLimit
;  = Set default state of Lock keys
;  = AHK Dark Mode
;  = Show/Hide OS files
;  = Initialise MultiClip
;  = Initialise hotstrings
;  = Tray Icon
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
;  = Move Mouse Pointer pixel by pixel
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
;  = Toggle Window Always-On-Top
;  = Process Priority
;    + SetPriority (enclosed Fn)
;  = Print Screen keys
; Launch Programs
;  = Hotkeys
;  = Hotstrings
; #HotIf Apps
;  = AHK windows
;    + AHK ToolTip
;    + AHK main window
;  = Calculator (classic)
;  = Firefox
;  = KeePass
;  = Notepad++
;    + Notepad++ main window
;    + Edit .ahk
;  = Dark Mode - Window Spy
;  = Windows Explorer
;    + Desktop + File Explorer
;      * Renaming
;      * Unselect
;      * Show folder/file size in ToolTip
;      * Delete Empty Folder
;      * Extract from folder & delete
;      * Copy full path
;      * Copy file names without path
;      * Copy file names without extension and path
;    + Desktop only
;      * Locate desktop background
;      * Find ahk_class
;    + File Explorer only
;      * Deletion Assist
;      * UnGroup
;      * Invert selection
;      * Horizontal Scrolling
;    + Run window (Win + R)
; #HotIf Groups
;  = Capitalise the first letter of a sentence
;  = Close With Esc/Q/W keys
;  = Horizontal Scrolling
;  = Media Keys Restored (disabled)
;  = Symbols In File Names keys
; Hotstrings - Actions
;  = MultiClip Hotstrings
;  = MultiClip testing
;  = Base64 Encode/Decode
;  = Date & Time
;    + Format Date / Time
;  = URL Encode/Decode
; Hotstrings - Text Replacement
;  = Find & Replace in Clipboard
;    + Find & Replace dot with space (StrReplace)
;    + Find & Replace dot with space (RegEx)
;  = Trim and change text
; Accented Vowels
; User-defined functions
;  = MyNotification Fn
;    + MyNotification_show
;    + MyNotification_close
;  = Catch error Fn
;    + CatchError_details
;    + CatchError_show
;  = AHK Dark Mode Fn
;    + ahkDarkMenu
;  = Toggle protected operating system (OS) files
;    + OSfiles_toggle
;    + OSfiles_check
;    + RefreshExplorer
;  = MultiClip Fn
;    + MultiClip_arrDefault
;    + MultiClip_check
;    + MultiClip_onChange
;    + MultiClip_addClip
;    + MultiClip_saveFileBackup
;    + MultiClip_vHotstrings
;    + MultiClip_pasteV
;    + MultiClip_cHotstrings
;    + MultiClip_pasteC
;    + MultiClip_pasteAll
;    + MultiClip_MenuFn
;    + MultiClip_pasteSlot
;  = Menu Icon Fn
;    + MenuIcons_add
;    + MenuIcons_error
;  = Paste instead of Send
;    + pasteThis
;    + pasteThis_clip
;  = Adjust Window Transparency
;    + WinTrans_get
;    + WinTrans_setMouse
;    + WinTrans_setMenu
;    + WinTrans_selection
;  = Change Text Case
;    + ChangeCase_menuFn
;    + ChangeCase_lower
;    + ChangeCase_Sentence
;    + ChangeCase_TitlE
;    + ChangeCase_UPPER
;    + ChangeCase_iNVERT
;    + ChangeCase_action
;  = Clipboard Fn
;    + Clip_wait
;    + Clip_call
;    + Clip_callVar
;  = ToolTip function
;    + ToolTipFn
;  = Wrap Text In Quotes or Symbols
;    + WrapText_menuFn
;    + WrapText_selection
;    + WrapText_action
;  = Base64 Encode/Decode
;    + Base64_encode
;    + Base64_decode
;    + Base64_decodeX2
;  = URL Encode/Decode
;    + URL_decode
;    + URL_encode
;  = MouseMove Fn
;    + MouseMove_screenCenter
;    + MouseMove_clientCenter
;  = Kill All Instances Of An App
;    + KillAll_titles
;  = Print Screen Fn
;    + ScreenSnip_menuFn
;    + ScreenSnip_selection
;    + ScreenSnip_printScreen
;    + ScreenSnip_fileOp
;  = Toggle Window Always-On-Top Fn
;    + AlwaysOnTop_enable
;    + AlwaysOnTop_disable
;  = Text correction by Menu
;    + TextCorrection_menuFn
;    + TextCorrection_action (enclosed Fn)
;  = File operations
;    + Folder_move
;  = String Fn
;    + CleanText_trim
;    + Str_LineLimit
;    + Str_LenLineimit
;  = Windows Explorer Fn
;    + Explorer_GetSize
;    + RBinQuery
;    + RBinVisible
;    + RBinHidden
;    + userSID
;    + checkDrivesFn
;    + RBinDisplay
;    + RBinGenerateTxt
;    + SizeFn
;    + DeleteEmptyFolder
;    + Explorer_ExtractFolder
;    + OpenFolder
;    + Loop_runShell (enclosed Fn)
;    + ExplorerTab_GetIdentity
;    + ExplorerTab_Navigate
;    + ExplorerTab_GetFolderPath
;    + ExplorerTab_GetSelectionPath
;    + Explorer_ValidatePath
;    + Explorer_FocusAddressBar (unused)
;    + WallpaperPath_v4
;    + nxtBackground
;  = File create, overwrite, append
;    + FileCreate_Overwrite
;  = MsgBox Fn
;    + ShowErrExit
;    + MsgBox_Custom
;    + MsgBox_ChangeButtonText
;    + MsgBox_setButtonNames (enclosed Fn)
;  = Compare Arrays & Remove Duplicates
;    + arrCompare_RemoveDuplicates
;    + arrCompare_error (enclosed Fn)
;    + arrCompare_results (enclosed Fn)
;  = #HotIf functions
;    + checkCalcView
;    + MarkletFn
;    + CloseIncrSearch
;  = Windows Registry
;    + RegJump
;    + Registry_ValidPath
;  = Control Panel Tools
;    + ControlPanel_menuFn
;    + ControlPanel_selection
;    + List of commands
; * Test

;------------------------------------------------------------------------------
; Settings
; Start of script code

#Requires AutoHotkey v2.0
#SingleInstance Force
#WinActivateForce
KeyHistory 500              ; Max 500

;------------------------------------------------------------------------------
; Auto-execute
; This section should always be at the top of your script

AHKname := "AHK v2 Showcase v3.00"

; Show notification with parameters
MyNotification_show("Loading " AHKname, 10000, 1550, 985, 1)
    ; - text
    ; - duration in milliseconds - 10000ms = 10 seconds
    ; - position on screen: x-axis 1550, y-axis 985 (bottom right corner) on 1920×1080 display resolution
    ; - end notification using sleep command (specify 0) or SetTimer (1) or manually (2)

loadingErrors := ""         ; store errors that occur during auto-execute

;--------
;  = Set Global variables

Global errorDetails := ""

;    + MenuShortcuts

; create array to store shortcuts used in Menu() functions
Global MenuShortcuts := StrSplit("1234567890QwertYuioPasdfG")

/* ; shortcuts for each menu item consist of numbers from number row, and letters from the rows below it in a QUERTY keyboard
; Customise the shortcut characters and their order by altering the characters in `MenuShortcuts` variable as needed
Menu Item 01 Shortcut 1 ; number row
Menu Item 02 Shortcut 2
Menu Item 03 Shortcut 3
Menu Item 04 Shortcut 4
Menu Item 05 Shortcut 5
Menu Item 06 Shortcut 6
Menu Item 07 Shortcut 7
Menu Item 08 Shortcut 8
Menu Item 09 Shortcut 9
Menu Item 10 Shortcut 0
Menu Item 11 Shortcut Q ; 1st letter row
Menu Item 12 Shortcut w
Menu Item 13 Shortcut e
Menu Item 14 Shortcut r
Menu Item 15 Shortcut t
Menu Item 16 Shortcut Y
Menu Item 17 Shortcut u
Menu Item 18 Shortcut i
Menu Item 19 Shortcut o
Menu Item 20 Shortcut P
Menu Item 21 Shortcut a ; 2nd letter row
Menu Item 22 Shortcut s
Menu Item 23 Shortcut d
Menu Item 24 Shortcut f
Menu Item 25 Shortcut G
; QYPG are capitals because selection underline causes confusion with other letters like o or v
; pressing shift + letter is not necessary because shortcuts are NOT case-sensitive
*/

;--------
;    + MenuIcons_path

; Icon Source: Calendar by Kalash - CC BY 4.0 - https://icon-icons.com/pack/Calendar/4173
; Icons were cropped using https://bulkimagecrop.com/ ; and converted to jpg and resized using mspaint (classic)
; To download and save the icons in below path, visit
; https://github.com/xypha/AHK-v2-scripts/blob/main/icons/Menu/

Global MenuIcons_path := A_ScriptDir "\icons\Menu\"
    ; icons are saved in default path of AHK's built-in variable: A_ScriptDir - The full path of the folder where the current script is located
    ; alternative location: A_MyDocuments - The full path to current user's "My Documents" folder. Usually corresponds to "C:\Users\<UserName>\Documents" (the final backslash is not included in the variable)
    ; modify this variable to match the location of icon files on your pc

;--------
;    + Send_LenLimit
; If length of text is short (less than or equal to 20 (user-defined)),
; use SendText mode (in the form of simulated keystrokes) instead of paste
; SendText: no attempt is made to translate characters (other than `r, `n, `t and `b) to keycodes

Global Send_LenLimit := 20
    ; change Send_LenLimit as needed, but note that bigger limit = longer execution time

;--------
;  = Set default state of Lock keys

; Turns on/off only once upon script load/reload
; Use "AlwaysOff" / "AlwaysOn" to force the key to stay On/Off permanently While script is running; in which case, add `Persistent` command to auto-execute section as needed

SetCapsLockState    "Off"   ; CapsLock   is off - Use SetCapsLockState "AlwaysOff" to force the key to stay off permanently, and add `Persistent` command if needed
SetNumLockState     "On"    ; NumLock    is ON
SetScrollLockState  "Off"   ; ScrollLock is off

;--------
;  = AHK Dark Mode

; check windows registry to see if dark mode is enabled
If not RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "AppsUseLightTheme", 1) {
    ; dark mode is enabled ; Boolean values: False = 0 and True = 1
    ; Global isLightMode := False                     ; 0 ; store RegRead results in variable for repeated use
    ahkDarkMenu()                                   ; enable dark theme for AHK menus
    }
; Else Global isLightMode := True                     ; 1 ; dark mode is NOT enabled

; download .ahk files from the `Lib` folder in this repo
; and save at the same under a `Lib` folder at the location of your .ahk script file

#Include "Lib\Dark Mode - ToolTip.ahk"              ; 2024.12.23
#Include <Dark Mode - MsgBox>                       ; 2024.12.23 ; different #Include format with same effect
; Dark Mode - Window Spy                              ; 2024.12.23

; manually comment out above lines if dark mode is NOT enabled because "#Include cannot be executed conditionally"
; disable dark mode commands "MyNotification_Gui.AddText" and "MyNotification_Gui.BackColor" in `MyNotification_show` function

;--------
;  = Show/Hide OS files

A_TrayMenu.Delete()                                 ; Delete standard menu
A_TrayMenu.Add("&Toggle OS files", OSfiles_toggle)  ; User-defined function
A_TrayMenu.Add()                                    ; Add a separator
A_TrayMenu.AddStandard()                            ; Restore standard menu
OSfiles_check()                                     ; Query registry and check/uncheck

;--------
;  = Initialise MultiClip

Global MultiClip_slotLimit := 25
    ; Limit the number of slots to 25 to match Win + V clipboard history slots ; customise limit to your needs
    ; Note:     Higher the number, higher the resource usage and slower the performance/response
    ; WARNING:  Removing the variable entirely may cause infinite loop / app hang

Global MultiClip_fileBackup := A_ScriptDir "\MultiClip_fileBackup.txt"
    ; txt file used instead of ini because 'values longer than 65,535 characters are likely to yield inconsistent results' when using IniRead, IniWrite commands

Global MultiClip_delim := Chr(0x007E) Chr(0x2022) Chr(0x007E)
    ; replace characters with Chr() command to prevent StrSplit() error when the above assigned delimiter itself is copied to clipboard
    ; 0x007E → ~ U+007E TILDE
    ; 0x2022 → • U+2022 BULLET : black small circle
    ; use a unique string because if an array-slot contains this delimiter by accident, saving and loading array from file will cause errors
    ; Recommendation: 3 or more characters, preferably symbols with one or more Unicode characters that are difficult to type on standard keyboard. For suggestions, look here - https://stackoverflow.com/questions/492090/least-used-delimiter-character-in-normal-text-ascii-128

Global MultiClip_arr := [] ; set Global variable and assign empty array

; Load array from file if file exists - inspired by https://www.autohotkey.com/boards/viewtopic.php?p=341809#p341809
; using `Try` command to prevent AutoHotkey from throwing error msg in case file is absent or not in correct path
Try MultiClip_arr := StrSplit(FileRead(MultiClip_fileBackup, "`n UTF-8"), MultiClip_delim, , MultiClip_slotLimit)
Catch as err { ; If error
    MultiClip_arr := MultiClip_arrDefault() ; load default values on start - 25 slots containing alphanumerical text
    loadingErrors .= CatchError_details(err, "MultiClip_arr[] default contents set!") "`n`n"
    }
Else MultiClip_check()

; run function whenever clipboard is changed such as when Ctrl + x (Cut) or Ctrl + c (Copy) is pressed
; or when clipboard is altered by other apps/programs
OnClipboardChange MultiClip_onChange

; add current clipboard contents to first clipboard slot in MultiClip_arr on start
MultiClip_addClip(A_Clipboard, 1) ; MultiClip_start = 1

; save `MultiClip_arr` contents to `MultiClip_fileBackup.txt` when the script exits
; except when it is killed by something like "End Task" via Taskbar, Task Manager or similar
OnExit MultiClip_saveFileBackup

;--------
;  = Initialise hotstrings

MultiClip_vHotstrings( MultiClip_slotLimit )    ; User-defined function creates serial hotstrings
MultiClip_cHotstrings( MultiClip_slotLimit )

;--------
;  = Tray Icon

path_to_TrayIcon := A_ScriptDir "\icons\Tray\1-512.ico"
    ; Icon source: https://www.iconsdb.com/caribbean-blue-icons/1-icon.html     ; CC License, see credits.md
    ; I like to number scripts 1, 2, 3... and link the scripts to Numpad shortcuts for easy editing -- see section on "Check & Reload AHK" below

Try TraySetIcon path_to_TrayIcon
Catch as err
    loadingErrors .= CatchError_details(err, "Path - " path_to_TrayIcon) "`n`n"

A_IconTip := A_ScriptName "`n" "AutoHotkey v" A_AhkVersion                              ; show AHK version in TrayTip

;--------
;  = Capitalise first letter opt-in Group

GroupAdd "CapitaliseFirstLetter_optIn", "ahk_class Notepad++"                           ; Notepad++ text editor
GroupAdd "CapitaliseFirstLetter_optIn", "ahk_class Notepad ahk_exe notepad.exe"         ; classic Notepad

;--------
;  = Close With Esc/Q/W Group

; GroupAdd "CloseWithEscQW"       , "ahk_exe DiskInfo64.exe"                              ; CrystalDiskInfo ; doesn't work because app starts with elevated privileges
; GroupAdd "CloseWithEscQW"       , "ahk_exe Taskmgr.exe"                                 ; Windows Task Manager ; requires UIAccess
GroupAdd "CloseWithEscQW"       , "Window Spy for AHKv2 ahk_class AutoHotkeyGUI"        ; AHK window spy
GroupAdd "CloseWithEscQW"       , "ahk_class CalcFrame"                                 ; classic calculator
GroupAdd "CloseWithEscQW"       , "Properties ahk_class #32770 ahk_exe mpc-hc.exe"      ; MediaInfo in MPC-HC x86
GroupAdd "CloseWithEscQW"       , "Properties ahk_class #32770 ahk_exe mpc-hc64.exe"    ; MediaInfo in MPC-HC x64
GroupAdd "CloseWithEscQW"       , "My Quota ahk_class TCustomForm ahk_exe networx.exe"  ; NetworX

;--------
;  = Horizontal Scrolling Group

GroupAdd "HorizontalScroll1"    , "ahk_class ApplicationFrameWindow"                    ; Modern UWP apps like calc and screen snip
GroupAdd "HorizontalScroll1"    , "ahk_class MozillaWindowClass"                        ; Firefox
GroupAdd "HorizontalScroll1"    , "ahk_class SALFRAME"                                  ; LibreOffice

;--------
;  = Media Keys Restored Group (disabled)
/* ; uncomment to use, if media keys are remapped to navigation keys in "Remap Keys" section

GroupAdd "MediaKeysRestored"    , "ahk_class MediaPlayerClassicW"                       ; MPC-HC
GroupAdd "MediaKeysRestored"    , "MPC-HC D3D Fullscreen"                               ; MPC-HC Full screen
GroupAdd "MediaKeysRestored"    , "ahk_class PotPlayer64"                               ; PotPlayer
; ahk_class RegEdit_RegEdit                                                               ; Registry Editor ; by default because AutoHotkey requires UIAccess

; To enable UIAccess for scripts by default,
; Navigate to the `UiAccessCommandKeyValue` in "HKEY_CLASSES_ROOT\AutoHotkeyScript\Shell\uiAccess\Command"
; copy it and paste it as (Default) value in "HKEY_CLASSES_ROOT\AutoHotkeyScript\Shell\Open\Command"
; Source: https://blog.danskingdom.com/Prevent-Admin-apps-from-blocking-AutoHotkey-by-using-UI-Access/
*/

;--------
;  = Symbols In File Names Group

GroupAdd "FileNameSymbols"      , "ahk_class CabinetWClass"                             ; Windows file explorer
GroupAdd "FileNameSymbols"      , "ahk_class EVERYTHING"                                ; Everything
GroupAdd "FileNameSymbols"      , "Renaming ahk_exe qbittorrent.exe"                    ; qBittorrent
GroupAdd "FileNameSymbols"      , "Save ahk_class #32770"                               ; Save As / Save File dialogue
GroupAdd "FileNameSymbols"      , "Export ahk_class #32770"
GroupAdd "FileNameSymbols"      , "Rename ahk_class #32770"

;--------
;  = WrapText variables

Global WrapText_vLead1 := "'"     , WrapText_vTrail1 := "'"                             ; single quotation '' - ' U+0027 : APOSTROPHE
Global WrapText_vLead2 := '"'     , WrapText_vTrail2 := '"'                             ; double quotation "" - " U+0022 : QUOTATION MARK
Global WrapText_vLead3 := "("     , WrapText_vTrail3 := ")"                             ; round brackets  ()
Global WrapText_vLead4 := "["     , WrapText_vTrail4 := "]"                             ; square brackets []
Global WrapText_vLead5 := "{"     , WrapText_vTrail5 := "}"                             ; flower brackets {}
Global WrapText_vLead6 := "``"    , WrapText_vTrail6 := "``"                            ; accent/backtick ``
Global WrapText_vLead7 := "%"     , WrapText_vTrail7 := "%"                             ; percent sign %%
Global WrapText_vLead8 := "‘"     , WrapText_vTrail8 := "’"                             ; ‘’ - ‘ U+2018 LEFT & ’ U+2019 RIGHT SINGLE QUOTATION MARK {single turned comma & comma quotation mark}
Global WrapText_vLead9 := "“"     , WrapText_vTrail9 := "”"                             ; “” - “ U+201C LEFT & ” U+201D RIGHT DOUBLE QUOTATION MARK {double turned comma & comma quotation mark}

;--------
;  = WrapText Disabled Group

GroupAdd "WrapText_disabled"      , "ahk_exe mpc-hc.exe"                                ; MPC-HC x86
GroupAdd "WrapText_disabled"      , "ahk_exe mpc-hc64.exe"                              ; MPC-HC x64
GroupAdd "WrapText_disabled"      , "ahk_class CalcFrame"                               ; Calculator (classic)
; Alt + number shortcuts are used to switch between Standard / Scientific / Programmer / Statistics calculators

; GroupAdd "WrapText_disabled"      , "ahk_class MSPaintApp"                              ; MS Paint (classic)
; commented out MSPaintApp from the group to enable wrapping text when inserting/editing text element (ClassNN: RICHEDIT50W1) - see `WrapText_action` for more info

;--------
;  = End auto-execute

; Reset notification timer to 1s after code in auto-execute section has finished running
SetTimer () => MyNotification_close(), -1000                                            ; 1s ; new thread

; show errors in MsgBox if any
If loadingErrors != ""
    SetTimer () => CatchError_show(loadingErrors), -1                                   ; 1ms ; new thread

; End auto-execute
Return

; Code can be placed anywhere in your script below this line

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

^!Numpad1:: {                                                   ; Ctrl + Alt + Numpad1 keys pressed together
MyNotification_show("Updating " AHKname,,, 985, 0)              ; 500ms ; use Sleep because reload cancels timers
Reload
Sleep 1000                                                      ; If successful, the reload will close this instance during the Sleep, so the line below will never be reached.
MsgBox "The script could not be reloaded!",, 16 + 262144        ; 16 IconX , 262144 Always-on-top
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
}                           ; disable keys - do nothing

;--------
;    + Keyboard keys

; Use Alt + Insert to toggle the 'Insert mode' and retain the key's function
; Note: ^Insert = Copy(^c) is Windows default behaviour and is not changed by this code
!Insert::Insert                         ; Source: https://gist.github.com/endolith/823381

LWin & Tab::AltTab                      ; Left Win key works as left Alt key - disables task view

RAlt::!Space                            ; Alt + Space brings up a window's title bar menu

^RCtrl::MButton                         ; press Left & Right Ctrl button to simulate mouse Middle Click

RCtrl & Up::        Send "{PgUp}"       ; Page up   - use "&" to create 2-key combo shortcut
RCtrl & Down::      Send "{PgDn}"       ; Page down - use a variable number of spaces before Send command without affecting the command itself
RControl & Left::   Send "{Home}"       ; Home      - use alternate key name for RCtrl
>^Right::           Send "{End}"        ; End       - use >^ instead of Right Ctrl button and skip using "&"

!m::WinMinimize "A"                     ; Alt+ M = Minimize active window
; PostMessage 0x0112, 0xF020,,, "A"       ; alternative, 0x0112 = WM_SYSCOMMAND, 0xF020 = SC_MINIMIZE

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

^CapsLock::^a                                   ; Select all
<#CapsLock::AltTab                              ; Switch windows with Right Win + CapsLock

;--------

+CapsLock:: {
SetCapsLockState "On"
MyNotification_show("CapsLock ON", 10000, 845)  ; 10000ms = 10s, change to match KeyWait timeout if needed
SetTimer () => CapsWait(), -1                   ; 1ms ; new thread
}

CapsWait() {
; runs in new thread and allows for quick toggling of CapsLock-state
; with +CapsLock / CapsLock / ESC keys in current thread

KeyWait "Esc", "d t10"                          ; hit ESC key to skip 10s timeout ; increase timeout duration to keep CapsLock ON for longer
SetCapsLockState "Off"                          ; Disables CapsLock immediately
MyNotification_close()                          ; and remove notification
}

;--------

CapsLock:: {                                    ; Turn off CapsLock immediately, if on
If GetKeyState("CapsLock", "T") {
    SetCapsLockState "Off"
    MyNotification_close()
    }
}

;------------------------------------------------------------------------------
;  = Move Mouse Pointer pixel by pixel
; Modified from http://www.computoredge.com/AutoHotkey/Downloads/MousePrecise.ahk

#Numpad1::MouseMove -1,  1, 0, "R"              ; Win + Numpad1 (SC04F) move down left    ↓←
#Numpad2::MouseMove  0,  1, 0, "R"              ; Win + Numpad2 (SC050) move down         ↓
#Numpad3::MouseMove  1,  1, 0, "R"              ; Win + Numpad3 (SC051) move down right   ↓→
#Numpad4::MouseMove -1,  0, 0, "R"              ; Win + Numpad4 (SC04B) move left         ←
#Numpad5::MouseMove_screenCenter()              ; Win + Numpad5 (SC04C) move center mouse • ; alternative - MouseMove_clientCenter()
#Numpad6::MouseMove  1,  0, 0, "R"              ; Win + Numpad6 (SC04D) move right        →
#Numpad7::MouseMove -1, -1, 0, "R"              ; Win + Numpad7 (SC047) move up left      ↑←
#Numpad8::MouseMove  0, -1, 0, "R"              ; Win + Numpad8 (SC048) move up           ↑
#Numpad9::MouseMove  1, -1, 0, "R"              ; Win + Numpad9 (SC049) move up right     ↑→

^#m::MouseMove 960,540                          ; Ctrl + Win + M = Test mouse position - screen centre in 1920 × 1080 display

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

winClass := WinGetClass(id)                                 ; Retrieves the specified window's class name
If (   winClass != "Shell_TrayWnd"                          ; exclude windows taskbar
    || winClass != "TopLevelWindowForOverflowXamlIsland"    ; System tray overflow window
    || winClass != "Windows.UI.Core.CoreWindow"             ; Notification Center
;   || winClass != "insert your app's window class"         ; uncomment to add more apps
    )
    WinClose(id)                                            ; sends a WM_CLOSE message to the target window
    ; PostMessage 0x0112, 0xF060,,, id ; alternative - same as pressing Alt+F4 or clicking a window's close button in its title bar
}

;--------
;    + WinKill with ^!F4
; Ctrl + Alt + F4 = attempt to kill window, applies to unresponsive ones when WinClose fails
; briefly tries to close the window normally and if that fails, attempts to terminate the window's process

^!F4::WinKill "A"

/* ; alternative
^!F4:: {
MouseGetPos ,, &id
ProcessClose WinGetProcessName(id)
}
*/

;--------
;    + Kill All Instances Of An App with ^!+F4
; Ctrl + Alt + Shift + F4 = attempt to Kill All Instances Of An App

^!+F4:: {
Process_Name := WinGetProcessName("A")

; DetectHiddenWindows False                                             ; default = only visible windows
vHWNDarr    := WinGetList("ahk_exe " Process_Name)
DetectHiddenWindows True                                                ; 1 ; get list of visible + invisible / hidden windows
ivHWNDarr   := WinGetList("ahk_exe " Process_Name)
DetectHiddenWindows False                                               ; 0 ; restore default

; remove already visible windows from this list
ivHWNDarr   := arrCompare_RemoveDuplicates(vHWNDarr, ivHWNDarr, 2)[2]   ; Returns [arr1, arr2, count := 0, errorMsg := 0]

vTitles     := KillAll_titles(vHWNDarr)                                 ; titles of visible windows

ivTitles    := KillAll_titles(ivHWNDarr)                                ; titles of hidden windows

Display := "Kill all instances of this app?"                                                                                    "`n`n"      ; `n = new line
        .  "Name of process:`t" Process_Name                                                                                    "`n"        ; `t = tab
        .  "Path of process:`t" WinGetProcessPath("ahk_exe " Process_Name)                                                      "`n`n"
        .  "No. of windows: " vHWNDarr.Length + ivHWNDarr.Length " (" vHWNDarr.Length " visible, " ivHWNDarr.Length " hidden)"  "`n"        ; no of windows ≠ no of processes
        .  "Titles of visible windows -"                                                                                        "`n"
        .  vTitles                                                                                                              "`n"
        .  "Titles of hidden windows:"                                                                                          "`n"
        .   ivTitles

myTitle := A_ScriptName " - Kill Process Windows - WARNING"
myText  := Str_LineLimit(Display)                                        ; limit to 35 lines

result_1 := MsgBox_Custom(myText, myTitle, 48 + 512, 3, "Yes", "Title List", "Cancel")
            ; 48 Icon! , 512 Default3 (make Cancel the default button to prevent accidental process kill)
            ; 3 = Yes / No / Cancel

If result_1 = "Yes" {                                                   ; MsgBox_Custom Button1 = Yes
    While ProcessExist(Process_Name)
        ProcessClose Process_Name
    ; alternative -
    ; Run A_ComSpec ' /C Taskkill /IM /F "' Process_Name '"'            ; alternative - you might see a flash of command prompt/terminal window. Brief explanation of flags -
    ; /C Carries out the command and then terminates
    ; /IM imagename
    ; /F forcefully terminates
    ; open Run dialogue (Win + R), paste "cmd.exe /?" and press OK, to see default flags
    ; then paste "Taskkill /?" (without the quotation marks) in cmd window and press enter to see 'Taskkill' specific flags, filters and examples
    }
Else If result_1 = "No" {                                               ; MsgBox_Custom Button2 = Title List
    filePath := A_Temp "\List of windows belonging to " Process_Name " - " FormatTime( , "yyyy-MM-dd @ HH：mm：ss") ".txt"
    FileCreate_Overwrite(filePath, Display)
    Run '"' filePath '"',,"Max"
    result_2 := MsgBox(myText, myTitle, 4 + 32 + 256 + 262144)          ; 4 YesNo , 32 Icon? , 256 Default2
    If result_2 = "Yes" {
        While ProcessExist(Process_Name)
            ProcessClose Process_Name
        }
    ; Else result = No                                                  ; do nothing further
    }
; Else result = Cancel                                                  ; MsgBox_Custom Button3 = Cancel ; do nothing further
}

/* More alternatives - untested - source: https://superuser.com/a/1220467/391770
> tasklist |find /i "sidebar"
sidebar.exe                  17252 Console                    1       209.680 K

> for /f "tokens=2" %A in ('tasklist ^|find /i "sidebar"') Do @Echo PID=%A
PID=17252

> for /f "tokens=2" %A in ('tasklist ^|find /i "sidebar"') Do @Taskkill /PID %A
*/

;------------------------------------------------------------------------------
;  = Adjust Window Transparency keys
; Modified from https://www.autohotkey.com/board/topic/667-transparent-windows/?p=148102

^+WheelUp:: {                                   ; increases Trans value, makes the window more opaque
MouseGetPos ,, &id
; id := WinExist("A")                           ; alternative - but 'Active' window might not always be the intended target
old_Trans := WinTrans_get(id)
If old_Trans < 255
    new_Trans := old_Trans + 20                 ; add 20, change for slower/faster transition
If old_Trans >= 255
    new_Trans := "Off"
WinTrans_setMouse(old_Trans, new_Trans, id)
}

^+WheelDown:: {                                 ; decreases Trans value, makes the window more transparent
MouseGetPos ,, &id
old_Trans := WinTrans_get(id)
If old_Trans > 20
    new_Trans := old_Trans - 20                 ; subtract 20, change for slower/faster transition
Else If old_Trans <= 20
    new_Trans := 1                              ; never set to zero (or negative), causes ERROR
WinTrans_setMouse(old_Trans, new_Trans, id)
}

F8::WinTrans_setMenu()

;------------------------------------------------------------------------------
;  = Recycle Bin shortcut

^Del:: {
If WinActive("Recycle Bin ahk_class CabinetWClass")             ; If windows file explorer is active and recycle bin is in the foreground, empty Bin
    FileRecycleEmpty
Else If WinExist("Recycle Bin ahk_class CabinetWClass")         ; If explorer is showing recycle bin but is in the background, activate it
    WinActivate

; use user defined function "OpenFolder" to
; (a) If an explorer is open but not showing recycle bin, change to Bin and
; (b) If an explorer is not open, then open Bin in explorer
Else {
    OpenFolder("{645ff040-5081-101b-9f08-00aa002f954e}")        ; comment out if not desired
    ; alternative to OpenFolder, directly open recycle bin in a new explorer window with below command
    ; Run "::{645ff040-5081-101b-9f08-00aa002f954e}"              ; If explorer is not open, then open Bin in explorer
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
If ThisHotkey = "^NumpadEnter" {
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
; Send "{Media_Stop}"                     ; stop playing all media
; DllCall("LockWorkStation")              ; Lock Workstation - source: https://gist.github.com/raveren/bac5196d2063665d2154#file-aio-ahk-L741

}

;--------
;  = Add Control Panel Tools to a Menu

#+x::ControlPanel_menuFn()   ; Win + Shift + x

;--------
;  = Change Text Case

!c::ChangeCase_menuFn()      ; Alt + C

;--------
;  = Wrap Text In Quotes or Symbols keys

#HotIf not WinActive("ahk_group WrapText_disabled")
; disables below hotkeys in apps that belonging to this group because they don't use it or have conflicts

!q::WrapText_menuFn()                                             ; Alt + Q

; WrapText Keys - Alt + Number (from the number row)
!1::WrapText_action( WrapText_vLead1 , WrapText_vTrail1 , 1)      ; enclose in single quotation '' - ' U+0027 : APOSTROPHE
!2::WrapText_action( WrapText_vLead2 , WrapText_vTrail2 , 1)      ; enclose in double quotation "" - " U+0022 : QUOTATION MARK
!3::WrapText_action( WrapText_vLead3 , WrapText_vTrail3 , 1)      ; enclose in round brackets  ()
!4::WrapText_action( WrapText_vLead4 , WrapText_vTrail4 , 1)      ; enclose in square brackets []
!5::WrapText_action( WrapText_vLead5 , WrapText_vTrail5 , 1)      ; enclose in flower brackets {}
!6::WrapText_action( WrapText_vLead6 , WrapText_vTrail6 , 1)      ; enclose in accent/backtick ``
!7::WrapText_action( WrapText_vLead7 , WrapText_vTrail7 , 1)      ; enclose in percent sign %%
!8::WrapText_action( WrapText_vLead8 , WrapText_vTrail8 , 1)      ; enclose in ‘’ - ‘ U+2018 LEFT & ’ U+2019 RIGHT SINGLE QUOTATION MARK {single turned comma & comma quotation mark}
!9::WrapText_action( WrapText_vLead9 , WrapText_vTrail9 , 1)      ; enclose in “” - “ U+201C LEFT & ” U+201D RIGHT DOUBLE QUOTATION MARK {double turned comma & comma quotation mark}
!0::WrapText_action( ""              , ""               , 1)      ; remove above quotes

#HotIf

;------------------------------------------------------------------------------
;  = Exchange adjacent letters
; place cursor between 2 letters. The letters reverse positions - `ab|c` becomes `ac|b`.
; Modified from http://www.computoredge.com/AutoHotkey/Downloads/LetterSwap.ahk

$!l:: {                                             ; Alt + L ; Test: AbC
Send "{Left}+{Right 2}"
clipVar := Clip_callVar(2)                          ; 2s, Exit
Send SubStr(clipVar,2) SubStr(clipVar,1,1) "{Left}"
}

;------------------------------------------------------------------------------
;  = Toggle Window Always-On-Top
; Modified from https://www.autohotkey.com/board/topic/94627-button-for-always-on-top/?p=596509

!t:: {                                              ; Alt + t

HWND        := WinGetID("A")
t           := WinGetTitle(HWND)
onTopTitle  := "! "                                 ; add prefix "! " to window title ; change as required
ExStyle     := WinGetExStyle(HWND)

If (ExStyle & 0x8)                                  ; 0x8 is WS_EX_TOPMOST
    AlwaysOnTop_disable(HWND, t, onTopTitle)
Else AlwaysOnTop_enable(HWND, t, onTopTitle)
}

;------------------------------------------------------------------------------
;  = Process Priority
; Hit `Win + P` to select and change the priority level of a process
; The current priority level of a process can be seen in the Windows Task Manager.

#p:: {
active_pid      := WinGetPID("A")
Process_Name    := WinGetProcessName("ahk_pid " active_pid)
PPGui           := Gui("AlwaysOnTop +Resize -MaximizeBox +MinSize240x230", "! Set Priority")

PPGui.AddText(, "Press ESCAPE to cancel.")
PPGui.AddText(, "Window (Top-most):"                            "`n"
            .   WinGetTitle("ahk_pid " active_pid)              "`n`n"
            .   "Process path:"                                 "`n"
            .   ProcessGetPath(active_pid)                      "`n`n"
            .   "Number of visible windows by the process: " WinGetList("ahk_pid " active_pid).Length    )
PPGui.AddText(, "Double-click to set a new priority level.")

LB := PPGui.AddListBox("r5 Choose1", ["Normal","High","Low","BelowNormal","AboveNormal"])
    ; Realtime omitted because any process not designed to run at Realtime priority might reduce system stability if set to that level ; add Realtime to ListBox if necessary
LB.OnEvent("DoubleClick", SetPriority)
PPGui.AddButton("default", "OK").OnEvent("Click", SetPriority)
PPGui.OnEvent("Escape"  , (*) => PPGui.Destroy())
PPGui.OnEvent("Close"   , (*) => PPGui.Destroy())
PPGui.Show

;--------
;    + SetPriority (enclosed Fn)

    SetPriority(*) {
    PPGui.Hide
    Try ProcessSetPriority(LB.Text, active_pid)
    Catch as err { ; if error
        info := "New priority could not be set - " LB.Text                      "`n"
             .  "Process: " Process_Name                                        "`n"
        SetTimer () => CatchError_show( CatchError_details(err, info) ), -1             ; 1ms ; new thread
        }
    Else { ; if successful
        msgText := ThisHotkey ":: Success! Priority changed!"                   "`n"
                .  "Process: " Process_Name                                     "`n"
                .  "Priority: " LB.Text
        ToolTipFn(msgText, 5000) ; 5s
        }
    PPGui.Destroy()
    }
}

;------------------------------------------------------------------------------
;  = Print Screen keys

; $PrintScreen::              ; keyboard hook $ ; commented out to preserve default function
#PrintScreen:: {            ; Win + PrintScreen
ScreenSnip_printScreen()    ; take screenshot, save and rename
}

; #+r::                       ; video snip shortcut, uncomment if desired
^PrintScreen::              ; Ctrl + Print Screen (key name = PrtSc, PrtScn or PrntScrn)
#+s:: {                     ; Win + Shift + s
ScreenSnip_menuFn()
}

;------------------------------------------------------------------------------
; Launch Programs

;  = Hotkeys

; Open Quick settings
~#a:: {
If WinWaitActive("Quick settings ahk_class ControlCenterWindow ahk_exe ShellHost.exe",, 3) ; 3s
    MouseMove 240, 575      ; move mouse to hover over Wi-Fi hotspot (top row, 2nd button on 1080p display)
    ; x-Axis = 240 (2nd button), change ± 140 for 1st or 3rd button respectively
    ; y-Axis 575 for 1st row, 700 for 2nd row
    ; use WheelDown key to access 3rd and 4th row
}

; Show hidden icons in tray
~#b:: {
If WinWaitActive("ahk_class Shell_TrayWnd",, 2) ; 2s
    Send "{Enter}"
    ; simulate {Enter} key press to open notification tray instead of just focusing on the ^ button
}

;------------------------------------------------------------------------------
;  = Hotstrings

; TextConverter AHK v1
:*x:chr?::Run A_ScriptDir "\Lib\TextConverter.ahk"

; KeyHistory AHK v1
:*x:khis+::Run A_ScriptDir "\Lib\KeyHistoryWindow.ahk"

; MouseHistroy AHK v1
:*x:mhis+::Run A_ScriptDir "\Lib\MouseHistoryWindow.ahk"

;------------------------------------------------------------------------------
; #HotIf Apps
; Tailor keyboard shortcuts, commands and functions to specific windows, apps or pre-defined groups of both

;  = AHK windows

;    + AHK ToolTip

#HotIf WinActive("ahk_class tooltips_class32")

^v:: {
If WinWaitClose("ahk_class tooltips_class32",, 10) ; 10s
    Send ThisHotkey
; Else Return ; do nothing
}

#HotIf

;--------
;    + AHK main window

#HotIf WinActive(".ahk - AutoHotkey v ahk_class AutoHotkey")
; because this is to enable below commands to apply on main windows of all running scrips irrespective of v1 or v2
; alternative to #HotIf WinActive(A_ScriptHwnd) applies only to the main window of current script

^Tab:: {                                                                    ; cycle through main window views like browser tabs

winID   := WinActive(".ahk - AutoHotkey v ahk_class AutoHotkey")
Text    := WinGetText()

If RegExMatch(Text, "^Script lines most recently executed") {               ; Found - ListLines
    If A_ScriptHwnd = winID
        ListVars                ; ListVars - Variables and their contents   ; ^v
    Else Send "^v"
    ToolTipFn("2 ListVars - Variables and their contents", 5000, 150, -25)  ; 5s ; add # to avoid confusion
    }
Else If RegExMatch(Text, "^Local Variables|^Global Variables") {            ; Found - ListVars
    If A_ScriptHwnd = winID
        ListHotkeys             ; ListHotkeys - Hotkeys and their methods   ; ^h
    Else Send "^h"
    ToolTipFn("3 ListHotkeys - Hotkeys and their methods", 5000, 150, -25)  ; 5s
    }
Else If RegExMatch(Text, "^Type\s+Off\?\s+Level\s+Running\s+Name") {        ; Found - ListHotkeys
    If A_ScriptHwnd = winID
        KeyHistory              ; KeyHistory - Key history and script info  ; ^k
    Else Send "^k"
    ToolTipFn("4 KeyHistory - Key history and script info", 5000, 150, -25) ; 5s
    }
Else If RegExMatch(Text, "^Window: ") {                                     ; Found - KeyHistory
    If A_ScriptHwnd = winID
        ListLines               ; ListLines - Lines most recently executed  ; ^L
    Else Send "^l"
    ToolTipFn("1 ListLines - Lines most recently executed", 5000, 150, -25) ; 5s
    }
Else {
    ToolTipFn(A_ThisHotkey ":: InStr Text not found!", 2000)                ; 2s
    Exit
    }
}

^+Tab:: {                                                                   ; cycle through main window views in reverse
winID   := WinActive(".ahk - AutoHotkey v ahk_class AutoHotkey")
Text    := WinGetText()

If RegExMatch(Text, "^Script lines most recently executed") {               ; Found - ListLines
    If A_ScriptHwnd = winID
        KeyHistory              ; KeyHistory - Key history and script info  ; ^k
    Else Send "^k"
    ToolTipFn("4 KeyHistory - Key history and script info", 5000, 150, -25) ; 5s
    }
Else If RegExMatch(Text, "^Local Variables|^Global Variables") {            ; Found - ListVars
    If A_ScriptHwnd = winID
        ListLines               ; ListLines - Lines most recently executed  ; ^L
    Else Send "^l"
    ToolTipFn("1 ListLines - Lines most recently executed", 5000, 150, -25) ; 5s
    }
Else If RegExMatch(Text, "^Type\s+Off\?\s+Level\s+Running\s+Name") {        ; Found - ListHotkeys
    If A_ScriptHwnd = winID
        ListVars                ; ListVars - Variables and their contents   ; ^v
    Else Send "^v"
    ToolTipFn("2 ListVars - Variables and their contents", 5000, 150, -25)  ; 5s
    }
Else If RegExMatch(Text, "^Window: ") {                                     ; Found - KeyHistory
    If A_ScriptHwnd = winID
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

!1::                                                                        ; Standard Calculator
!2::                                                                        ; Scientific
!3::                                                                        ; Programmer
!4::                                                                        ; Statistics
{
If ControlGetClassNN(ControlGetFocus()) = "Edit1" && ThisHotkey = "!3"      ; focus is on editing history and Alt + 3 is pressed
    WrapText_action( WrapText_vLead3 , WrapText_vTrail3 )                   ; enclose in round brackets ()
Else Send ThisHotkey "^{F4}"                                                ; switch calculator type and go to basic view
}

; Toggle Date Calculation view
^d:: {
If checkCalcView() = 1
    Send "^{F4}"                ; Ctrl + F4 = Basic view
Else Send "^e"                  ; Ctrl + E  = Date Calculation view
}

; Toggle Unit conversion view
^u:: {
If checkCalcView() = 2
    Send "^{F4}"                ; Ctrl + F4 = Basic view
Else Send "^u"                  ; Ctrl + U  = Unit conversion view
}

^w::^F4                         ; Go to Basic view

^Tab:: {                        ; cycle through calculator views like browser tabs

CalcView := checkCalcView()

If CalcView = 0                 ; Basic view
    Send "^u"                   ; Ctrl + U  = Unit conversion view
Else If CalcView = 1            ; Unit conversion view
    Send "^e"                   ; Ctrl + E  = Date Calculation view
Else If CalcView = 2            ; Date Calculation view
    MenuSelect , , "View", "Worksheets", "Mortgage"
Else If CalcView = 3            ; Worksheets > Mortgage
    MenuSelect , , "View", "Worksheets", "Vehicle lease"
Else If CalcView = 4            ; Worksheets > Vehicle lease
    MenuSelect , , "View", "Worksheets", "Fuel economy (mpg)"
Else If CalcView = 5            ; Worksheets > Fuel economy (mpg)
    MenuSelect , , "View", "Worksheets", "Fuel economy (L/100 km)"
Else                            ; Worksheets > Fuel economy (L/100 km)
    Send "^{F4}"                ; Ctrl + F4 = Basic view
}

^+Tab:: {

CalcView := checkCalcView()

If CalcView = 0                 ; Basic view
    MenuSelect , , "View", "    Worksheets", "Fuel economy (L/100 km)"
Else If CalcView = 1            ; Unit conversion view
    Send "^{F4}"                ; Ctrl + F4 = Basic view
Else If CalcView = 2            ; Date Calculation view
    Send "^u"                   ; Ctrl + U  = Unit conversion view
Else If CalcView = 3            ; Worksheets > Mortgage
    Send "^e"                   ; Ctrl + E  = Date Calculation view
Else If CalcView = 4            ; Worksheets > Vehicle lease
    MenuSelect , , "View", "Worksheets", "Mortgage"
Else If CalcView = 5            ; Worksheets > Fuel economy (mpg)
    MenuSelect , , "View", "Worksheets", "Vehicle lease"
Else                            ; Worksheets > Fuel economy (L/100 km)
    MenuSelect , , "View", "Worksheets", "Fuel economy (mpg)"
}

#HotIf

;--------
;  = Firefox
; for custom shortcuts in Firefox, Try add-on: https://addons.mozilla.org/en-US/firefox/addon/shortkeys/

#HotIf WinActive("ahk_class MozillaWindowClass")
    ; main window ; excludes other dialogue boxes like "Save As" originating from ahk_exe firefox.exe

; Ctrl + Shift + F = close Find Bar
^+f::Send "^f{Esc}"

; Alt + H = open Homepage
!h::Send "^w^t"
    ; Background image loads correctly. Go backwards and Go forwards button history is lost, but not permanently. Use ^+t to restore most recent closed tab and check tab history OR use check recent history in Firefox Library.
    ; Send "^Labout:home{Enter}" ; alternative - Go backwards and Go forwards button history is preserved, but blank grey background may be seen instead of new tab background image

; Ctrl + Shift + O = open library / bookmark manager
^+o:: {
If WinActive(" — Mozilla Firefox")  ; If not new tab, then open new one
    Send "^t"
Else Send "^l"                      ; If new tab, focus address bar
Sleep 250                           ; 250ms ; wait for focus - change as per your system performance
Send "^a"                           ; select all to replace with pasted text
pasteThis("chrome://browser/content/places/places.xhtml")
Send "{Enter}"
}

; Run saved bookmarklet via keyword
; Check my bookmarklet repo - https://github.com/xypha/Bookmarklets
^m:: {                              ; Ctrl + M
If WinActive(" IMDb")
    MarkletFn("imdblink")           ; keyword for `Open IMDb trailer in a new tab`
Else If WinActive(".pdf")
    MarkletFn("pdfdark")            ; keyword for "PDF dark"
Else If WinActive("Preferences - Invidious")
    MarkletFn("invpref")            ; keyword for `Set Invidious preferences in two clicks`
Else Send ThisHotkey
}

; Ctrl + N sends Ctrl + T to open new Tab instead of new window
; ^n::^t

; Ctrl + Shift + Q = Exit (Disable default Firefox shortcut)
^+q::Return

#HotIf

;------------------------------------------------------------------------------
;  = KeePass

#HotIf WinActive("ahk_exe KeePass.exe")

^n::^i          ; Ctrl + N = new key entry, not new database

#HotIf

;------------------------------------------------------------------------------
;  = Notepad++

;    + Notepad++ main window

#HotIf WinActive("ahk_class Notepad++")

^d::^f          ; disable duplicate line (SCI_SELECTIONDUPLICATE), open find

; close incremental search if focused when using below shortcuts
^g::            ; Ctrl + G                          = open Go To dialogue box
^s::            ; Ctrl + S                          = save / save all
^+s::           ; Ctrl + Shift + S                  = save / save all
^w::            ; Ctrl + W                          = close tab
^h::            ; Ctrl + H                          = open Find-Replace dialogue box
^+h::           ; Ctrl + Shift + H                  = open Open from file history…
^+f::           ; Ctrl + Shift + F                  = open Find dialogue box - updated 2024.11.07
^Tab::          ; Ctrl + Tab                        = switch tab forwards
^+Tab::         ; Ctrl + Shift + Tab                = switch tab backwards
{
CloseIncrSearch(ThisHotkey)
}

^t::^n          ; open new document like a browser Tab instead of SCI_LINETRANSPOSE
~^+t::Return    ; !! added this to prevent error with above hotkey for some reason !!

#HotIf

;--------
;    + Edit .ahk
; my custom AHK lexer - https://gist.github.com/xypha/baf6167058b4c37066c8257530c0ffad

#HotIf WinActive(".ahk - Notepad++ ahk_class Notepad++")

; correct symbols
:*:<=::<=   ; less-or-equal
:*:=<::<=

:*:>=::>=   ; greater-or-equal
:*:=>::>=

; capitalise commands for clarity
:*:catch::Catch
:*:continue::Continue
:*:else::Else
:*:exit::Exit
:*:false::False
:*:fileRecycle::FileRecycle
:*:finally::Finally
:*:global::Global
:*:gui::Gui
:*:listlines::ListLines
:*:loop::Loop
:*:mouseclick::MouseClick
:*:raw::Raw
:*:return::Return
:*:sleep::Sleep
:*:static::Static
:*:thisHotkey::ThisHotkey
:*:true::True
:*:try::Try
:*:while::While
:*:winactivate::WinActivate
:*:winclose::WinClose
:*:winexist::WinExist
::break::Break          ; disable immediate replacement `*` to avoid conflict with breakdown
::msgbox::MsgBox
::off::Off
::run::Run              ; not * conflict with (Run)ning
::tooltip::ToolTip

; capitalise keys for clarity
:*:click::Click
:*:end::End
:*:home::Home
:*:left::Left
:*:right::Right
:*:space::Space
::BS::Backspace
::down::Down            ; not * conflict with (Down)town
::enter::Enter
::tab::Tab              ; not * conflict with (Tab)ular
::up::Up                ; not * conflict with (Up)town

; expand commonly used inbuilt commands
:*:editpaste::EditPaste "String", "Control", "WinTitle"
:*:filedelete::FileRecycle "file_pattern*"
:*:menuselect::MenuSelect "WinTitle",, "Search", "Everything"
:*:mousemove::MouseMove x, y,, "R" `; client, relative
:*:msgbox+::MsgBox("",, 262144) `; 262144 Always-on-top
:*:run+::Run '"exe_path" "' variable '"',,"Max"
:*:settimer::SetTimer () => , -1 `; 1ms `; new thread
:*:trim+::Trim(msgText, "``r``n``t``s``v``f")
:*:winactive::WinActive("Title"){Left 2}
:*:winmove::WinMove x, y, w, h `; screen
:*:winwait::WinWaitActive("Title",, 2) `; 2s

; expand commonly used custom commands
:*:clip+::A_Clipboard
:*:index+::A_Index
:*:send ::Send ""{Left}             ; Send ""

; long section break
:*x:;-+:: {
pasteThis(";------------------------------------------------------------------------------")
Send "{Enter}"
}

; short section break
:*x:;--::   Send ";--------{Enter}"

; section headings
:*x:sec1+:: SendText ";  = "       ; SendText instead of `T` option, so that `+` ≠ Shift and trailing Space
:*x:sec2+:: SendText ";    + "
:*x:sec3+:: SendText ";      * "

; comment block open / close
:*:/*::
:*:*/:: {
Send "/*{Enter 2}`*/{Up}"
}

; using different types of continuation sections
:*x:cont1+::Send '" `;" continuation{Enter}({Enter}only_text{Enter})`"{Del} `;"'

:*x:cont2+:: {
v := ' ; continuation section
(
 ( `; continuation section
"text" "``n" function "``n")
)'
pasteThis(v)
}

#HotIf

;------------------------------------------------------------------------------
;  = Dark Mode - Window Spy

#HotIf WinActive("Window Spy for AHKv2 ahk_class AutoHotkeyGUI")

^c:: {
A_Clipboard := StrReplace(Clip_callVar(2), "`r`n", "`s") ; 2s, Exit ; copy contents to one line
}

#HotIf

;------------------------------------------------------------------------------
;  = Windows Explorer

;    + Desktop + File Explorer

#HotIf WinActive("ahk_class Progman") || WinActive("ahk_class CabinetWClass")

;      * Renaming

F1::F2      ; rename instead of accidental opening help in MS edge
F3::F2      ; rename instead of accidental focus on search, use default Ctrl + F to focus on search instead

;--------
;      * Unselect
; Unselect multiple files/folders
; Source: https://superuser.com/questions/78891/is-there-a-keyboard-shortcut-to-unselect-in-windows-explorer

^+a::F5     ; Ctrl + Shift + A = unselect by sending {F5} key ; same as Right Click > Refresh
; alternative - WindowsRefreshOrRun()

;--------
;      * Show folder/file size in ToolTip

^+s:: {

If WinActive("Recycle Bin - File Explorer ahk_class CabinetWClass ahk_exe explorer.exe") {
    ; obtain drive letters, Breakdown per drive, total no of files, total size of files, list of files
    ; and display results in MsgBox
    RBinDisplay(RBinQuery())
    Exit
    }
; Else ; proceed with below code

; get pathContent and pathType = (1) multiple folder/files, (2) single folder, (3) single file, (4) CLSID or (0) invalid mono/multi-line
pathContent := ExplorerTab_GetSelectionPath()
pathType    := Explorer_ValidatePath(pathContent)

If pathType = 4 || pathType = 0                                             ; (4) CLSID or (0) invalid
    ShowErrExit(A_ThisHotkey ":: ExplorerTab_GetFolderPath() extracted path is CLSID or invalid!" "`n"
            .   pathContent                                                                            )

; calculate folder size and display
result := Explorer_GetSize(pathType, pathContent)                           ; Return [SizeB, errorDetails, pathContent, files, fileCount]

; check size and display results in tooltip
If result[2] != ""                                                          ; errorDetails present
    ShowErrExit(result[5] "files with size of " SizeFn(result[1])   "`n"    ; fileCount and SizeB
            .   "Errors:"                                           "`n"
            .   result[2]                                           "`n"    ; errorDetails
            .   "Path Content:"                                     "`n"
            .   result[3]                                               )   ; pathContent

Else If result[1] = 0                                                       ; SizeB zero and no errors
    ToolTipFn( Str_LenLineimit(result[3]) "`n" "Empty folder/file!", 3000)  ; 3s ; show pathContent with line length 100 and 30 lines

Else {                                                                      ; SizeB ≠ 0
    info := Str_LenLineimit(result[3])    "`n"                              ; line length 100 and 30 lines ; pathContent
         .  result[5] " files"           "`n"                               ; fileCount
         .  "Size: " SizeFn(result[1])                                      ; SizeB

    ToolTipFn( info , 4000)                                                 ; 4s
    }
}

;--------
;      * Delete Empty Folder

^+d:: {                                             ; Ctrl + Shift + D
pathContent := ExplorerTab_GetSelectionPath()       ; Returns pathContent

msgText :=  ThisHotkey ":: ExplorerTab_GetSelectionPath() → DeleteEmptyFolder()"  "`n"
        .   DeleteEmptyFolder(pathContent)          ; Returns success / error info

MsgBox Str_LineLimit(msgText),, 262144              ; limit lines to 35 ; 262144 Always-on-top
}

;--------
;      * Extract from folder & delete

^+e:: {

; get pathContent and pathType = (1) multiple folder/files, (2) single folder, (3) single file, (4) CLSID or (0) invalid
path        := ExplorerTab_GetSelectionPath()
pathType    := Explorer_ValidatePath(path)

If pathType != 1 && pathType != 2
    ShowErrExit(ThisHotkey ":: ExplorerTab_GetSelectionPath() did not extract a valid folder!"  "`n"
            .   "Path Type: " pathType " with path content:"                                    "`n"
            .   path                                                                                )
; Else pathType = 1 or 2

; store Return info for later display
msgText := ""

; Run extraction function
Loop Parse path, "`n", "`r"
    msgText .= Explorer_ExtractFolder(A_LoopField)
    ; Returns info from SourceFolder error / DestinationFolder error / extraction_errors / DeleteEmptyFolder() aborted, success, error info

; show msgs after extraction
MsgBox Str_LineLimit(msgText),, 262144 ; limit lines to 35 ; 262144 Always-on-top
}

;--------
;      * Copy full path
; Modified from https://www.autohotkey.com/boards/viewtopic.php?p=61084#p61084
; Example of output: C:\Program Files\Mozilla Firefox\firefox.exe

^+c:: {                         ; Ctrl + Shift + C
Clip_call(2)                    ; 2s, Exit
A_Clipboard := A_Clipboard      ; change to plain text
}

;--------
;      * Copy file names without path
; Example of output: firefox.exe

!n:: {                                                      ; Alt + N
A_Clipboard := RegExReplace(Clip_callVar(2), "\w:\\|.+\\")  ; 2s, Exit ; remove path
}

;--------
;      * Copy file names without extension and path
; Example of output: firefox

^!n:: { ; Ctrl + Alt + N
files := RegExReplace(Clip_callVar(2), "\w:\\|.+\\")        ; 2s, Exit ; remove path
files := RegExReplace(files, "\.[\w]+(`r`n|`n)","`n")       ; remove ext, CR
A_Clipboard := RegExReplace(files, "\.[\w]+$")              ; remove last ext
}

#HotIf

;------------------------------------------------------------------------------
;    + Desktop only

#HotIf WinActive("Program Manager ahk_class Progman ahk_exe explorer.exe") || WinActive("ahk_class Shell_TrayWnd ahk_exe explorer.exe")

;      * Locate desktop background

#w:: { ; win + w

; If mouse is on desktop, but WinActive is Tray, MouseClick to focus on desktop
MouseGetPos , , &id
If WinGetClass(id) = "Progman" && WinGetClass("A") = "Shell_TrayWnd" {
    MouseMove_screenCenter()
    MouseClick
    }

path    := WallpaperPath_v4(ThisHotkey)

result  := MsgBox_Custom(path, "Current wallpaper location", 256, 3, "In Explorer?", "Next Pic?", "OK")
           ; 3 = Yes / No / Cancel ; 256 Button2 default

; Actions
If result = "Yes" ; Button1 renamed to "In Explorer?"
    ; open path in windows file explorer and select file
    Run 'explorer.exe /n,/e,/select, "' path '"'

Else If result = "No" { ; Button2 renamed to "Next Pic?"

    ; ask slideshow to go to next wallpaper
    nxtBackground()

    ; delete old wallpaper
    Try FileRecycle path
    Catch as err {          ; If error, open photo in default app
        info := CatchError_details(err, "WallpaperPath_v4() FileRecycle failure - " path) "`n`n"
        Try Run path
        Catch as err {      ; If Run error, copy path to clipboard
            info .= CatchError_details(err, "WallpaperPath_v4() Run failure")
            SetTimer () => CatchError_show(info), -1   ; 1ms ; new thread
            Return
            }
        Else                ; Run executed without error
            SetTimer () => CatchError_show(info), -1   ; 1ms ; new thread
        }
    }
; Else result = "Cancel" ; do nothing ; Button3 renamed to "OK"
}

#HotIf

;--------
;      * Find ahk_class

#w:: {
UniqueID := WinActive("A")

MsgBox(ThisHotkey "::WallpaperPath_v4() not executed! WinActive is not ahk_class Progman."  "`n"
    .  "Title: "    WinGetTitle(UniqueID)                                                   "`n"
    .  "ahk_class " WinGetClass(UniqueID)                                                   "`n"
    .  "ahk_exe "   WinGetProcessName(UniqueID)                                                 ,, 262144) ; 262144 Always-on-top
}

;------------------------------------------------------------------------------
;    + File Explorer only

#HotIf WinActive("ahk_class CabinetWClass")

;      * Deletion Assist
; change folder/file selection to the next folder/file when the currently focused folder/file is deleted

NumpadDel::
Del:: {                                 ; trigger when delete or Numpad delete button are pressed
ClassNN := ControlGetClassNN(ControlGetFocus())
If ClassNN = "SysTreeView321" {         ; stop delete in navigation pane
    ToolTipFn("Focus is in Navigation Pane! Delete Aborted!", 2000) ; 2s
    ControlFocus "DirectUIHWND2", "File Explorer ahk_class CabinetWClass"
    }
Else If ClassNN ~= "DirectUIHWND"       ; use RegEx match to check if focus is on file list
    Send "{" ThisHotkey "}{Down}"       ; send down arrow to select next item
Else Send "{" ThisHotkey "}"
}

;--------
;      * UnGroup
; Change the annoying `Group by Date modified` default in Downloads folder to `Group by Date (None)`

^g:: {                                  ; Ctrl + G
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

#HotIf

;--------
;    + Run window (Win + R)

#HotIf WinActive("Run ahk_class #32770 ahk_exe explorer.exe")

:*b0x:startup+:: {
Send "{Esc}" ; close dialogue
Run "shell:startup"
; OpenFolder(A_Startup)
}

; Jump to Registry key
:*b0x:regJump+:: {
Send "{Esc}"
RegJump()
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
; ~？::          ; Unicode U+003F QUESTION MARK
; Triggers       ; add or disable one or more as needed
{
cfc1 := InputHook("L1 V C","{Space}{LShift}{RShift}{CapsLock}", "a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,x,y,z") ; captures 1st character, visible, case sensitive ; .a → .A
cfc1.Start
cfc1.Wait
If cfc1.EndReason = "Match" {
    If ThisHotkey = "~." || ThisHotkey = "~NumpadDot"
    ; in case NumpadDot or . is the trigger, then don't capitalise because typing the website address and file names is problematic ;  Example.exe → Example.exe (no change)
        Send "{Backspace}" cfc1.Input

    ; Else If ThisHotkey = "~NumpadEnter" || ThisHotkey = "~Enter"
    ; in case NumpadEnter or Enter is the trigger, capitalise 1st character BUT don't add space, because space is not necessary when creating a new paragraph
    ; commented out because trigger is disabled
    ;     Send "{Backspace}+" cfc1.Input

    Else { ; in case ! or ? is the trigger, then add a space and capitalise 1st character ; !a → ! A  and ?b → ? B
        Send "{Backspace} +" cfc1.Input
        ; SoundBeep 1500, 50 ; play a sound when successful - Frequency(a number between 37 and 32767), Duration in milliseconds
        ; SoundPlay "C:\Windows\Media\Windows Information Bar.wav" ; alternative to SoundBeep
        ; SoundPlay A_WinDir "\Media\Windows Balloon.wav"          ; alternative to SoundBeep
        }
    Exit ; stop capture if match found
    }
If cfc1.EndKey = "Space" { ; prevent cfc2 from firing for numbers or symbols. Example: 0.2ms is not changed to 0.2Ms
    cfc2 := InputHook("L1 V C","{Space}{LShift}{RShift}{CapsLock}", "a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,x,y,z") ; captures 2nd character, visible, case sensitive ; . a → . A
    cfc2.Start
    cfc2.Wait
    If cfc2.EndReason = "Match"
        Send "{Backspace}+" cfc2.Input
    }
}

#HotIf

; several other AHK v1 auto-capitalisation scripts are good, such as the one linked above
; and one from computoredge - http://www.computoredge.com/AutoHotkey/Downloads/AutoSentenceCap.ahk
; and many others that use different methods to achieve this goal. Try a few and see what works for you.

;------------------------------------------------------------------------------
;  = Close With Esc/Q/W keys

#HotIf WinActive("ahk_group CloseWithEscQW")

Esc::WinClose "A"   ; sends a WM_CLOSE message to the target window

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

/* ; uncomment to use, if media keys are remapped to navigation keys in "Remap Keys" section
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
; replace \/:*?"<>| with ＼⧸ ： ✲ ？＂＜＞｜
; inspired by the file renaming scheme of "yt-dlg" app - https://oleksis.github.io/youtube-dl-gui/
; comment out the ones you don't desire

#HotIf WinActive("ahk_group FileNameSymbols")

; extra character `+` to trigger replacement in order to prevent False positive match while entering RegEx
:?*:\+::{U+FF3C}                    ; \ → ＼ | replace U+005C REVERSE SOLIDUS : backslash            → U+FF3C FULLWIDTH REVERSE SOLIDUS
:?*:|+::{U+FF5C}                    ; | → ｜ | replace U+007C VERTICAL LINE : vertical bar, pipe     → U+FF5C FULLWIDTH VERTICAL LINE
:?*:`:+::{U+FF1A}                   ; : → ：  | replace U+003A COLON                                  → U+FF1A FULLWIDTH COLON
:?*:*+::{U+2732}                    ; * → ✲ | replace U+002A ASTERISK : star                         → U+2732 OPEN CENTRE ASTERISK

:?*:/::{U+29F8}                     ; / → ⧸  | replace U+002F SOLIDUS : slash, forward slash, virgule → U+29F8 BIG SOLIDUS
:?*:?::{U+FF1F}                     ; ? → ？ | replace U+003F QUESTION MARK                          → U+FF1F FULLWIDTH QUESTION MARK
:?*:"::{U+FF02}                     ; "; → ＂| replace U+0022 QUOTATION MARK : double quote          → U+FF02 FULLWIDTH QUOTATION MARK
:?*:<::{U+FF1C}                     ; < → ＜ | replace U+003C LESS-THAN SIGN                         → U+FF1C FULLWIDTH LESS-THAN SIGN
:?*:>::{U+FF1E}                     ; > → ＞ | replace U+003E GREATER-THAN SIGN                      → U+FF1E FULLWIDTH GREATER-THAN SIGN

; Template -
; :*:*::{U+}                     ; ? → ? | replace ?     → ?

#HotIf

;------------------------------------------------------------------------------
; Hotstrings - Actions

;  = MultiClip Hotstrings

:?*x:v0+::MultiClip_pasteV(10)          ; same as v10+ ; pastes value in slot #10 in MultiClip_arr ; default value "j10"

:?*xR:c1+::Send MultiClip_arr[1]        ; R = Raw mode ; Send first entry in raw mode, useful when Ctrl + V is disabled such as on banking sites

:?*x:c0+::MultiClip_pasteC(10)          ; same as c10+

; !v::                                    ; Alt + v
:?*x:c++:: {                            ; c++
MultiClip_MenuFn( MultiClip_pasteSlot ) ; show MultiClip_Menu
}

;--------
;  = MultiClip testing

:*:testclip+:: {

; save current array contents to file
MultiClip_saveFileBackup()
    ; If script is reloaded after test, saved file will be deleted. In this case, restore array contents by
    ; (a) Exiting this script (b) restoring deleted file from recycle bin (c) running this script again

A_Clipboard := "Slot 1 Shortcut 1"
Global MultiClip_arr := MultiClip_arrDefault()
MultiClip_MenuFn( MultiClip_pasteSlot )     ; show menu - MultiClip_Menu
}

; restore old MultiClip_arr contents from file after testing or from previous file when needed
:*x:restoreclip+:: {

; enable local default delimiter if restoring MultiClip_arr after changing Global default delimiter
; MultiClip_delim := Chr(0x007E) Chr(0x2022) Chr(0x007E)

Global MultiClip_arr := ""
MultiClip_arr := StrSplit(FileRead(MultiClip_fileBackup, "`n UTF-8"), MultiClip_delim, , MultiClip_slotLimit)   ; restore from file
MultiClip_MenuFn( MultiClip_pasteSlot )     ; show menu - MultiClip_Menu
}

;------------------------------------------------------------------------------
;  = Base64 Encode/Decode

:*x:b64+::pasteThis(Base64_encode(A_Clipboard))     ; Encoding

:*x:b1-::pasteThis(Base64_decode(A_Clipboard))      ; Decoding

:*x:b2-::pasteThis(Base64_decodeX2(A_Clipboard))    ; Decoding twice

;------------------------------------------------------------------------------
;  = Date & Time

;    + Format Date / Time

:*x:d++::       Send FormatTime(, "yyyy.MM.dd")             ; sends 2021.02.31
:*x:date+::     Send FormatTime(, "dd.MM.yyyy")             ; sends 28.03.2020
:*x:time+::     Send FormatTime(, "h:mm tt")                ; sends 6:48 PM
:*x:datetime+:: Send FormatTime(, "dd.MM.yyyy h:mm tt")     ; sends 28.03.2020 6:46 PM

;------------------------------------------------------------------------------
;  = URL Encode/Decode

:*x:url+::pasteThis(URL_encode(A_Clipboard))

/* Encode URL
    Example: https://www.google.com/
    Copy example URL to clipboard
    Triger `URL_encode` function by typing `url+`
    Output: https%3A%2F%2Fwww.google.com%2F
*/

:*x:url-::pasteThis(URL_decode(A_Clipboard))

/* Decode URL
    Example: https%3A%2F%2Fwww.google.com%2F
    Copy example URL to clipboard
    Triger `URL_decode` function by typing `url-`
    Output: https://www.google.com/
*/

;------------------------------------------------------------------------------
; Hotstrings - Text Replacement

;  = Find & Replace in Clipboard

;    + Find & Replace dot with space (StrReplace)

:*:.++:: { ; hotstring ".++"
A_Clipboard := StrReplace(A_Clipboard,"."," ") ; replace *each* dot with space
}

/*
Find text:          "ABC..def.GHI"
Replacement text:   "ABC  def GHI"
*/

;--------
;    + Find & Replace dot with space (RegEx)

:*:.r++:: { ; hotstring ".r++"
A_Clipboard := RegExReplace(A_Clipboard,"\.+"," ") ; replace *one or more* dots with single space
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
myText      := RegExReplace(A_Clipboard, "\s+", A_Space)    ; replace one or more \s with A_Space (which = " " single space)
A_Clipboard := RegExReplace(myText, "^ +| +$")              ; trim leading/trailing spaces
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
myText      := RegExReplace(A_Clipboard, "m)^ +| +$")       ; m) = multi-line, trim leading/trailing spaces
myText      := RegExReplace(myText, "^\s+|\s+$|`r")         ; trim \r, leading/trailing \n
A_Clipboard := RegExReplace(myText, " +", A_Space)          ; trim double spaces
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
; Accented Vowels

:*x:a'+::TextCorrection_menuFn("À , à , Á , á , Â , â , Ã , ã , Ä , ä")
:*x:e'+::TextCorrection_menuFn("È , è , É , é , Ê , ê , Ñ , ñ , Ë , ë")
:*x:i'+::TextCorrection_menuFn("Ì , ì , Í , í , Î , î , Õ , õ , Ï , ï")
:*x:o'+::TextCorrection_menuFn("Ò , ò , Ó , ó , Ô , ô ,         Ö , ö")
:*x:u'+::TextCorrection_menuFn("Ù , ù , Ú , ú , Û , û ,         Ü , ü")
:*x:y'+::TextCorrection_menuFn(        "Ý , ý ,                 Ÿ , ÿ")

/* Source: https://sites.psu.edu/symbolcodes/windows/codealt/
1 Grave Capital
2 Grave Lower Case
3 Acute Capital
4 Acute Lower Case
5 Circumflex Capital
6 Circumflex Lower Case
7 Tilde Capital
8 Tilde Lower Case
9 Umlaut Capital
0 Umlaut Lower Case

1   2   3   4   5   6   7   8   9   0
À   à   Á   á   Â   â   Ã   ã   Ä   ä
È   è   É   é   Ê   ê   Ñ   ñ   Ë   ë
Ì   ì   Í   í   Î   î   Õ   õ   Ï   ï
Ò   ò   Ó   ó   Ô   ô           Ö   ö
Ù   ù   Ú   ú   Û   û           Ü   ü
        Ý   ý                   Ÿ   ÿ
*/

;------------------------------------------------------------------------------
; User-defined functions

;  = MyNotification Fn
; search for `ToolTipFn` for alternative

;    + MyNotification_show

MyNotification_show(myText, myDuration := 500, xAxis := 1550, yAxis := 985, endType := 1) { ; 500ms
    ; xAxis x yAxis based on display resolution of 1920 × 1080

Dec         := 18       ; decrease yAxis by this number per line
MaxLen      := 40       ; max characters per line
    ; These two variables should be adjusted based on font defined below by `MyNotification_Gui.SetFont` command
    ; and width defined under `MyNotification_Gui.AddText` command and (maybe?) other factors like DPI

; use RegExReplace to breakdown text to `MaxLen` limit and
; CleanText_trim() remove leading/trailing Spaces, `r, double Spaces and trim
needle      := "m)^(?=.{" MaxLen -2 ",})(.*?)([ \.\?!\-]+|$)(.*)$"
myText      := CleanText_trim(RegExReplace(myText, needle, "$1…`n…$2"))

If StrReplace(myText, "`n" , "`n",, &lineCount)
    yAxis   := yAxis - (Dec * lineCount)

Global MyNotification_Gui := Gui("+AlwaysOnTop -Caption +ToolWindow")
    ; Global so that MyNotification_close() works
    ; -Caption disables title bar and thick window border/edge which is present by default
    ; +ToolWindow avoids a taskbar button and an Alt-Tab menu item for this Gui()

MyNotification_Gui.SetFont("s9 w1000", "Arial")
    ; s9 font Size 9 points, w1000 Weight/boldness (1 - 1000, 400 is normal, 700 is bold)

MyNotification_Gui.BackColor := "2C2C2E"
    ; "2C2C2E" for dark mode background - remove this command for light mode

MyNotification_Gui.AddText("cWhite w230 Left vMyNotifText", myText)
    ; "cWhite" font Colour for dark mode text - remove this word for light mode
    ; w230 Width accommodates about 40 characters
    ; Left-justifies the control's text within its available width
    ; -Wrap disables word-wrapping (enabled by default)

MyNotification_Gui.Show("x" xAxis " y" yAxis " NoActivate")
    ; NoActivate avoids deactivating the currently active window

If endType = 0 {                                                    ; use Sleep to trigger MyNotification_close
    Sleep Abs(myDuration)                                           ; always positive number
    MyNotification_close()
    }
Else If endType = 1                                                 ; use SetTimer to trigger MyNotification_close
    SetTimer () => MyNotification_close(), Abs(myDuration) * -1       ; new thread ; Abs() * -1 for always negative number
; Else ; closer = 2                                                 ; trigger MyNotification_close manually by adding command to the script where necessary
}

;--------
;    + MyNotification_close

MyNotification_close() {
Try MyNotification_Gui.Destroy()
}

;------------------------------------------------------------------------------
;  = Catch error Fn

;    + CatchError_details

CatchError_details(err, info := "", errorType := Type(err)) {

If A_ThisHotkey
    Details := A_ThisHotkey ":: "
Else Details := "No hotkey/hotstring :: "

Details .=  err.What                                            "`n"
        .   "Type: " errorType  " @ Line: " err.Line            "`n"

If info
    Details .=  "User Info: " info                              "`n"

; add [err] properties to details

If A_LastError != 0
    Details .=  "OSError"   OSError(A_LastError).Message        "`n"
If err.Message
    Details .=  "Message: " err.Message                         "`n"
If err.Extra
    Details .=  "Extra: "   err.Extra                           "`n"

; add AHK help file information to details

If err.What = "RegRead" {
    If errorType = "OSError"
        Details .=  "Help Info: non-existent key or value / permission error"           "`n"
    Else  ; Error or other?
        Details .=  "Help Info: unknown - check help file"                              "`n"
    }

Else If err.What = "ControlGetHwnd" {
    If errorType = "TargetError"
        Details .=  "Help Info: Window or control could not be found"                   "`n"
    Else  ; Error or other?
        Details .=  "Help Info: unknown - check help file"                              "`n"
    }

Else If err.What = "WinSetAlwaysOnTop" {
    If errorType = "TargetError"
        Details .=  "Help Info: target not found"                                       "`n"
    Else If errorType = "OSError"
        Details .=  "Help Info: failed to set OnTop"                                    "`n"
    Else  ; Error or other?
        Details .=  "Help Info: unknown - check help file"                              "`n"
    }

Else If err.What = "FileRead" {
    If errorType = "OSError"
        Details .=  "Help Info: opening or reading the file"                            "`n"
    Else If errorType = "MemoryError"
        Details .=  "Help Info: file > 4 GB / unable to allocate enough memory"         "`n"
    Else  ; Error or other?
        Details .=  "Help Info: unknown - check help file"                              "`n"
    }

Else If err.What = "MenuSelect" {
    If errorType = "TargetError"
        Details .=  "Help Info: Target not found / no Win32 menu"                       "`n"
    Else If errorType = "ValueError"
        Details .=  "Help Info: Item not found / not final menu parameter"              "`n"
    Else  ; Error or other?
        Details .=  "Help Info: unknown - check help file"                              "`n"
    }

Else If err.What = "EditPaste" {
    If errorType = "TargetError"
        Details .=  "Help Info: Window or control could not be found"                   "`n"
    Else If errorType = "OSError"
        Details .=  "Help Info: Message could not be sent to the control"               "`n"
    Else  ; Error or other?
        Details .=  "Help Info: unknown - check help file"                              "`n"
    }

Else If err.What = "DllCall" {
    If errorType = "OSError"
        Details .=  "Help Info: HRESULT or something else?"                             "`n"
    Else If errorType = "TypeError"
        Details .=  "Help Info: Unexpected type / not a string or positive integer"     "`n"
    Else If errorType = "ValueError"
        Details .=  "Help Info: return/arg type error"                                  "`n"
    Else  ; Error or other?
        Details .=  "Help Info: multiple possibilities - check help file"               "`n"
    }

Else Details .=  "Help Info: not in CatchError_details() - check help file"             "`n"

Return Details
}

; SetTimer () => CatchError_show( CatchError_details(err, info) ), -1     ; 1ms ; new thread

;--------
;    + CatchError_show

CatchError_show(Details) {

Details := Trim(Details, "`r`n`t`s`v`f")
msgText := Str_LineLimit(Details)                                       ; limit lines to 35
msgOpt  := 0                                                            ; default (OK) button in MsgBox

If msgText != Details {
    fileName    := A_ScriptDir "\CatchError_details.txt"
    Details     := "`n`n" FormatTime( , "yyyy-MM-dd @ HH：mm：ss")
                .  "`n"   Details
    FileAppend( Details , fileName , "`n UTF-8")                        ; create new file
    msgText .=  "`n`n" "Check below file for full details"  "`n"
            .   fileName                                    "`n`n"
            .   "View file?"
    msgOpt  := 4                                                        ; 4 YesNo button in MsgBox
    }

result := MsgBox(msgText, A_ScriptName " - Error!", msgOpt + 262144)    ; 16 IconX , 262144 Always-on-top
If result = "Yes"
    Run fileName
}

;------------------------------------------------------------------------------
;  = AHK Dark Mode Fn

;    + ahkDarkMenu
; primary source: https://stackoverflow.com/a/58547831/894589
; and modified by lexikos - https://www.autohotkey.com/boards/viewtopic.php?p=426482#p426482
; and modified by mcd     - https://www.autohotkey.com/boards/viewtopic.php?p=511756#p511756

ahkDarkMenu() {
Static uxtheme              := DllCall("GetModuleHandle", "str", "uxtheme", "ptr")
Static SetPreferredAppMode  := DllCall("GetProcAddress", "ptr", uxtheme, "ptr", 135, "ptr")
Static FlushMenuThemes      := DllCall("GetProcAddress", "ptr", uxtheme, "ptr", 136, "ptr")

DllCall(SetPreferredAppMode, "int", 1) ; 0 = Default, 1 = AllowDark, 2 = ForceDark, 3 = ForceLight, 4=Max
DllCall(FlushMenuThemes)
}

;------------------------------------------------------------------------------
;  = Toggle protected operating system (OS) files
; inspiration from https://www.autohotkey.com/board/topic/82603-toggle-hidden-files-system-files-and-file-extensions/?p=670182

;    + OSfiles_toggle
; alternative - Run ToggleSystemFiles.bat as administrator to toggle settings - https://superuser.com/a/1151851/391770

OSfiles_toggle(*) { ; optional variables (ItemName, ItemPos, MyMenu) if called from A_TrayMenu

If show_protected_files = 0 { ; enable if disabled
    Try RegWrite "1", "REG_DWORD", "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSuperHidden"
    Catch as err {
        SetTimer () => CatchError_show( CatchError_details(err, "OSfiles_toggle 0 → 1 (enabling) failed!") ), -1   ; 1ms ; new thread
        Exit
        }
    OSfiles_check()
    SetTimer () => RefreshExplorer(), -1         ; 1ms ; new thread
    }
Else { ; disable if enabled
    Try RegWrite "0", "REG_DWORD", "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSuperHidden"
    Catch as err {
        SetTimer () => CatchError_show( CatchError_details(err, "OSfiles_toggle 1 → 0 (disabling) failed!") ), -1   ; 1ms ; new thread
        Exit
        }
    OSfiles_check()
    SetTimer () => RefreshExplorer(), -1         ; 1ms ; new thread
    }
}

;--------
;    + OSfiles_check

OSfiles_check() { ; tray tick mark
If not IsSet(show_protected_files)
    Global show_protected_files := RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSuperHidden")
If show_protected_files = 0
    A_TrayMenu.UnCheck "&Toggle OS files"
Else A_TrayMenu.Check "&Toggle OS files"
}

;--------
;    + RefreshExplorer
; modified from https://www.autohotkey.com/boards/viewtopic.php?p=482766#p482766

RefreshExplorer() {
Explorers := ComObject("Shell.Application").Windows
Explorers.Item(ComValue(0x13, 0x8)).Refresh()   ; 0x13 = VT_UI4 32-bit unsigned int, 0x8 = SWC_DESKTOP
For Window in Explorers
    ; If (Window.Name != "Internet Explorer")   ; not necessary in Win11 and later
    Window.Refresh()
; Else ; If for Loop fails with zero iterations, open new explorer window because one doesn't already exist ; uncomment this and below command if desired
;     Run 'explorer.exe',,"Max"

If WinExist("ahk_class CabinetWClass") && not WinActive() ; If Windows File Explorer window exists and not active
    WinActivate
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

*/

;------------------------------------------------------------------------------
;  = MultiClip Fn
; Sources - AutoHotkey.chm::/docs/lib/OnClipboardChange.htm#ExBasic
; and https://www.autohotkey.com/boards/viewtopic.php?p=326827#p326827
; and MultiClip v1 - https://www.autohotkey.com/boards/viewtopic.php?p=332658#p332658
; MultiClip v1 used or modified AHK v1 code from https://autohotkey.com/board/topic/4567-clipstep-step-through-multiple-clipboards-using-ctrl-x-c-v/
; and https://geekdrop.com/content/super-handy-autohotkey-ahk-script-to-change-the-case-of-text-in-line-or-wrap-text-in-quotes

;    + MultiClip_arrDefault

MultiClip_arrDefault(number := MultiClip_slotLimit) {

arrDefault := []

Loop number
    arrDefault.Push("Default Slot 01 Shortcut " MenuShortcuts[A_Index])

Return arrDefault
}

;--------
;    + MultiClip_check

MultiClip_check(number := MultiClip_slotLimit) {
Loop number {
    Try MultiClip_arr.Get(A_Index)      ; .Get throws an IndexError if Index is zero or out of range
    Catch                               ; add index and default value
        MultiClip_arr.InsertAt(A_Index, "Default Slot " A_Index " Shortcut " MenuShortcuts[A_Index] )
    Else                                ; Index is valid
        If MultiClip_arr[A_Index] = ""  ; value is blank, assign default value
            MultiClip_arr[A_Index] := "Default Slot " A_Index " Shortcut " MenuShortcuts[A_Index]
        ; Else ; value is not blank, do nothing and Continue
    }
}

;--------
;    + MultiClip_onChange

MultiClip_onChange(DataType) {

If DataType = 0 {                                                       ; 0 Clipboard is now empty
    ; ToolTipFn("Clipboard DataType: 0 - Clipboard is now empty", 1000)   ; 1s
    Return
    }

Else If DataType = 2 {                                                  ; 2 Clipboard contains something entirely non-text such as a picture
    ToolTipFn("Clipboard DataType: 2 - Non-text copied", 1000)          ; 1s
    Return
    }

; Else DataType = 1 ; Clipboard contains text (including files copied from Windows File Explorer)
; check and add to MultiClip_arr. In case of files, file path is copied to MultiClip_arr
Else MultiClip_addClip(A_Clipboard)
}

;--------
;    + MultiClip_addClip

MultiClip_addClip(text, MultiClip_start := 0) {             ; MultiClip_start = default False

tmp := StrReplace(text, "`r`n", "`n")                       ; fix for SendInput sending Windows line-breaks

If RegExMatch(tmp,"^\s+$") {                                ; don't insert empty strings into MultiClip_arr
    If MultiClip_start != 1                                 ; If NOT on startup/reload
        ToolTipFn("[~ Only \s ~]", 2000)                    ; 2s ; show alert instead
    ; Else ; MultiClip_start = 1 , so don't show ToolTip
    Exit
    }

tmp := Trim(tmp, "`r`n`t`s`v`f")
; RegExReplace(tmp,"^\s*|\s*$")                             ; remove leading/trailing \s = [\r\n\t\f\v]

; use Loop to check if tmp is already in an array. If found, remove it and retrieve its `Index`
Loop MultiClip_slotLimit {
    If tmp == MultiClip_arr[A_Index] {                      ; case-sensitive
        MultiClip_arr.RemoveAt(A_Index)
        FoundClip := A_Index
        Break
        }
    }
MultiClip_arr.InsertAt(1, tmp)                              ; insert current clipboard contents in the first slot
MultiClip_arr.Length := MultiClip_slotLimit                 ; reset number of slots to previously defined limit
If IsSet(FoundClip) && FoundClip = 1 && MultiClip_start = 1
    Return
Else ToolTipFn( Str_LenLineimit(MultiClip_arr[1]) , 1000)   ; 1s  ; line length 100 and 30 lines
}

;--------
;    + MultiClip_saveFileBackup

; use * as parameter because OnExit Callback accepts two parameters
; OnExit MyCallback(ExitReason, ExitCode)

MultiClip_saveFileBackup(*) {
Result := ""

; save current MultiClip_arr contents to variable
Loop (MultiClip_arr.Length > MultiClip_slotLimit ? MultiClip_slotLimit : MultiClip_arr.Length)
    Result .= MultiClip_arr[A_Index] MultiClip_delim

; remove trailing MultiClip_delim to prevent MultiClip_arr.Length from exceeding MultiClip_slotLimit on restoration of MultiClip_arr
Result := SubStr(Result, 1, StrLen(MultiClip_delim) * -1)

Try FileRecycle MultiClip_fileBackup                     ; send old file to recycle bin if one exists
    ; old clipboard contents can be retrieved by restoring MultiClip_fileBackup from recycle bin
    ; alternatively, use `FileDelete` command to delete permanently
; Catch ; useless as script terminates if the user chooses Exit from the tray menu or main window menu, or the script is asked to terminate as a result of Reload or #SingleInstance
FileAppend Result, MultiClip_fileBackup, "`n UTF-8"      ; create new file and save current MultiClip_arr contents
}

;--------
;    + MultiClip_vHotstrings

MultiClip_vHotstrings(number) {
Loop number
    Hotstring(":?*x:v" A_Index "+", MultiClip_pasteV)
}

/* use Loop to replace serialised hotstrings
:?*:v1+::
.
.
.
:?*:v20+:: {
PasteV(ThisHotkey)
}
*/

;--------
;    + MultiClip_pasteV

MultiClip_pasteV(hk) {                  ; MyCallback(HotstringName := ThisHotkey)
RegExMatch(hk, "\d+", &SubPat)
pasteThis( MultiClip_arr[SubPat[]] )
; Try Send MultiClip_arr[SubPat[]]      ; alternative
}

;--------
;    + MultiClip_cHotstrings

MultiClip_cHotstrings(number) {
Loop number {
    If A_Index != 1
        Hotstring(":?*x:c" A_Index "+", MultiClip_pasteC)
    ; Else ; A_Index = 1, do nothing because c1+ hotstring already assigned :?*xR:c1+::Send MultiClip_arr[1]
    }
}

/* use Loop to replace serialised hotstrings
:?*:c0+:: ; same as c10
:?*:c2+::
.
.
.
:?*:c20+:: {
PasteC(ThisHotkey)
}
*/

;--------
;    + MultiClip_pasteC

MultiClip_pasteC(hk) {                  ; MyCallback(HotstringName := ThisHotkey)
RegExMatch(hk, "\d+", &SubPat)
MultiClip_pasteAll(SubPat[])
}

;--------
;    + MultiClip_pasteAll

MultiClip_pasteAll(hkey := MultiClip_arr.Length) {
results := ""

Loop hkey
    results .= MultiClip_arr[A_Index] "`n"

pasteThis( RTrim(results, "`r`n`t`s`v`f") ) ; remove trailing LF
}

;--------
;    + MultiClip_MenuFn

MultiClip_MenuFn(FnName) {

Global MultiClip_Menu := Menu()     ; starts building a pop-up menu
MultiClip_Menu.Delete()             ; retained because slotContent may have changed

LoopNo := MultiClip_arr.Length > MultiClip_slotLimit ? MultiClip_slotLimit : MultiClip_arr.Length

; add menu items
Loop LoopNo {

    ; separator for replacing `n (new line) when displaying menu items - U+2385 ⎅ WHITE SQUARE WITH CENTRE VERTICAL LINE : center
    ; find alternative symbols using BabelMap at https://www.babelstone.co.uk/
    ; retain only first 100 characters
    slotContent := SubStr( StrReplace(MultiClip_arr[A_Index], "`n", " " Chr(0x2385) " ") , 1, 100)

    MultiClip_Menu.Add("&" MenuShortcuts[A_Index] "  → " slotContent, FnName)
    ; When the menu is displayed, a slot can be selected by pressing the key corresponding to the character preceded by the ampersand (&)
    ; These selection shortcuts correspond to the number/alphabet/symbol before `→` and are obtained from ClipShortcuts array
    ; When the menu is displayed, shortcuts are usually underlined, but sometimes don't appear when some symbols are used
    }

; add menu icons
MenuIcons_add(MultiClip_Menu, LoopNo)

; show pop-up menu
MultiClip_Menu.Show
}

;--------
;    + MultiClip_pasteSlot

MultiClip_pasteSlot(ItemName, ItemPos, MyMenu) {
pasteThis(MultiClip_arr[ItemPos])
}

;------------------------------------------------------------------------------
;  = Menu Icon Fn

;    + MenuIcons_add

MenuIcons_add(MenuFn, LoopNo) {
Loop LoopNo {

    Try MenuFn.SetIcon(A_Index "&", MenuIcons_path A_Index "-20.jpg", , 20)
    ; `Try` command is used to prevent AutoHotkey from throwing error msgs in case icon files are absent or not in correct path.
    ; WARNING: Using icons in menu may slow performance.
    ; A slight delay between menu request and display may be noticeably present on some systems (especially in low-end ones like mine; probably to resize/rescale icons).
    ; This is normal and expected. Comment out the `ClipMenu.SetIcon` line if this is not acceptable.
    ; Default size is 16. Increased to 20 because icon numbers are not clear. Please let me know if you find icons that look better with `ahkDarkMenu` by creating an Issue

    Catch as err {     ; If error, add menu item indicating that menu icons are missing
        MenuFn.Add("&/  → Menu icons are missing! Press / to check icon path", MenuIcons_error)
        Break
        }
    }
}

;--------
;    + MenuIcons_error

MenuIcons_error(ItemName, ItemPos, MyMenu) {
mytext := "Menu icons were not loaded for: " MyMenu                                     "`n"
       .  "Icons should be saved in below path -"                                       "`n"
       .   MenuIcons_path                                                               "`n`n"
       .  "Check path to see if icon files exist and are correctly named."              "`n"
       .  "If not, modify `MenuIcons_path` variable to the path containing menu icons"  "`n"
       .  "To download menu icons, visit https://github.com/xypha/AHK-v2-scripts/blob/main/icons/Menu/"
MsgBox mytext,, 262144 ; 262144 Always-on-top
}

;------------------------------------------------------------------------------
;  = Paste instead of Send
; Modified from https://www.autohotkey.com/boards/viewtopic.php?p=483549#p483549 and https://www.autohotkey.com/boards/viewtopic.php?p=483588#p483588
; alternative to inbuilt command - EditPaste String, Control [, WinTitle, WinText, ExcludeTitle, ExcludeText]

;    + pasteThis

pasteThis(pasteText) {
If Trim(pasteText, "`r`n`t`s`v`f") = "" {
    ToolTipFn(A_ThisHotkey ":: pasteThis(pasteText) is empty!", 5000) ; 5s
    Return
    }

If pasteText == A_Clipboard {               ; case-sensitive
    Send "^v"
    Return
    }

If StrLen(pasteText) <= Send_LenLimit       ; default = 20 ; If short text, Send keystrokes instead of paste
    SendText pasteText                      ; text mode to prevent unintended key press when text contains '^+!#{}'
Else pasteThis_clip(pasteText)
}

;--------
;    + pasteThis_clip

pasteThis_clip(pasteText) {
OnClipboardChange MultiClip_onChange, 0     ; disable callback
tmp         := ClipboardAll()
A_Clipboard := pasteText
Send "^v"                                   ; paste
Sleep 500                                   ; 500ms wait for paste command to execute before restoring clipboard
A_Clipboard := tmp
OnClipboardChange MultiClip_onChange, 1
}

;------------------------------------------------------------------------------
;  = Adjust Window Transparency

;    + WinTrans_get

WinTrans_get(id) {
Trans := WinGetTransparent(id)
If not Trans
    Return 255
Else Return Trans
}

;--------
;    + WinTrans_setMouse

WinTrans_setMouse(old_Trans, new_Trans, id) {
If new_Trans = "Off"
    WinSetTransparent 255, id
    ; Set transparency to 255 before using Off - might avoid window redrawing problems such as a black background
    ; If the window still fails to be redrawn correctly, try WinRedraw, WinMove or WinHide + WinShow for a possible workaround
WinSetTransparent new_Trans, id
ToolTipFn("Transparency: " old_Trans " → " new_Trans, 2000) ; 2s
}

;--------
;    + WinTrans_setMenu
; modified from http://www.computoredge.com/AutoHotkey/Downloads/Always_on_Top.ahk

WinTrans_setMenu() {
MouseGetPos ,, &WinID               ; identify window id
; WinID := WinExist("A")              ; alternative - but 'Active' window might not always be the intended target
Global WinID                        ; so that WinTrans_selection can use it to set transparency

If IsSet(WinTrans_menu) {
    WinTrans_menu.Show
    Return
    }
; Else ; Continue as below

Global WinTrans_menu := Menu()      ; starts building a pop-up menu
; WinTrans_menu.Delete()              ; replaced by If IsSet()

; add menu items
WinTrans_menu.Add("&1  → 255 Opaque"            , WinTrans_selection )
WinTrans_menu.Add("&2  → 190 Translucent"       , WinTrans_selection ) ; Semi-opaque
WinTrans_menu.Add("&3  → 125 Semi-transparent"  , WinTrans_selection )
WinTrans_menu.Add("&4  →  65 Nearly Invisible"  , WinTrans_selection )
WinTrans_menu.Add("&5  →   1 Invisible"         , WinTrans_selection ) ; never set to zero, causes ERROR

; add menu icons
MenuIcons_add(WinTrans_menu, 5)

; show pop-up menu
WinTrans_menu.Show
}

;--------
;    + WinTrans_selection

WinTrans_selection(ItemName, ItemPos, MyMenu) {
old_Trans := WinGetTransparent(WinID)
If old_Trans = "" && WinActive(WinID)
    old_Trans := 255

new_Trans := Trim(SubStr(ItemName, 7, 3))
WinSetTransparent new_Trans, WinID
If new_Trans = 255
    WinSetTransparent "Off", WinID ; Specifying Off - may improve performance and reduce usage of system resources
ToolTipFn("Transparency: " old_Trans " → " Trim(SubStr(ItemName, 7)), 2000) ; 2s
}

;------------------------------------------------------------------------------
;  = Change Text Case
; inspired by https://geekdrop.com/content/super-handy-AutoHotkey-ahk-script-to-change-the-case-of-text-in-line-or-wrap-text-in-quotes

;    + ChangeCase_menuFn

ChangeCase_menuFn() {
If IsSet(ChangeCase_menu) {
    ChangeCase_menu.Show
    Return
    }
; Else ; Continue as below

Global ChangeCase_menu := Menu()     ; starts building a pop-up menu
; ChangeCase_menu.Delete()           ; replaced by If IsSet()

; add menu items
ChangeCase_menu.Add("&1  → lower case"      , ChangeCase_lower    )
ChangeCase_menu.Add("&2  → Sentence case"   , ChangeCase_Sentence )
ChangeCase_menu.Add("&3  → Title Case"      , ChangeCase_TitlE    )
ChangeCase_menu.Add("&4  → UPPER CASE"      , ChangeCase_UPPER    )
ChangeCase_menu.Add("&5  → iNVERT cASE"     , ChangeCase_iNVERT   )

; add menu icons
MenuIcons_add(ChangeCase_menu, 5)

; show pop-up menu
ChangeCase_menu.Show
}

;--------
;    + ChangeCase_lower

ChangeCase_lower(*) {
Clip_call(2)                                    ; 2s, Exit
ChangeCase_action(StrLower(A_Clipboard))
}

;--------
;    + ChangeCase_Sentence

ChangeCase_Sentence(*) {
Clip_call(2)                                    ; 2s, Exit
ChangeCase_action(RegExReplace(StrLower(A_Clipboard), "(((^\s*|([.!?]+\s*))[a-z])|\Wi\W)", "$U1")) ; Code Credit #1
}

; RegEx explanation -
; \s = [\r\n\t\f\v]
; $U1 = back reference uppercase 1
; \W = [^a-zA-Z0-9_] = any character that is NOT alphabet, number, underscore

;--------
;    + ChangeCase_TitlE

ChangeCase_TitlE(*) {
Clip_call(2)                                    ; 2s, Exit
ChangeCase_action(StrTitle(A_Clipboard))
}

;--------
;    + ChangeCase_UPPER

ChangeCase_UPPER(*) {
Clip_call(2)                                    ; 2s, Exit
ChangeCase_action(StrUpper(A_Clipboard))
}

;--------
;    + ChangeCase_iNVERT

ChangeCase_iNVERT(*) {
Clip_call(2)                                    ; 2s, Exit
inverted := ""
Loop Parse A_Clipboard {                        ; Code Credit #2
    If StrLower(A_LoopField) == A_LoopField     ; case-sensitive ; * Code Credit #3
        inverted .= StrUpper(A_LoopField)       ; *
    Else inverted .= StrLower(A_LoopField)      ; *
    }
ChangeCase_action(inverted)
}

; Unicode TestString  :=
; abcdefghijklmnopqrstuvwxyzéâäàåçêëèïîìæôöòûùÿáíóúñABCDEFGHIJKLMNOPQRSTUVWXYZÉÂÄÀÅÇÊËÈÏÎÌÆÔÖÒÛÙŸÁÍÓÚÑ
; Unicode iNVERT cASE :=
; ABCDEFGHIJKLMNOPQRSTUVWXYZÉÂÄÀÅÇÊËÈÏÎÌÆÔÖÒÛÙŸÁÍÓÚÑabcdefghijklmnopqrstuvwxyzéâäàåçêëèïîìæôöòûùÿáíóúñ

;--------
;    + ChangeCase_action

ChangeCase_action(caseText) {
string := StrReplace(caseText, "`r")    ; remove \r
Len := StrLen(string)
pasteThis(string)                       ; Paste
If Len <= Send_LenLimit                 ; and select text string if ≤ Send_LenLimit
    Send "+{Left " Len "}"              ; change `Send_LenLimit` as needed, bigger Len = longer execution time
}

; Code Credit #1 - NeedleRegEx pattern modified from https://www.autohotkey.com/board/topic/24431-convert-text-uppercase-lowercase-capitalized-or-inverted/?p=158295
; Code Credit #2 - idea for loop modified from https://www.autohotkey.com/boards/viewtopic.php?p=58417#p58417
; Code Credit #3 - 3 lines of code with a comment "; *" were adapted from a (inaccurate) answer generated from a query to KudoAI's DuckDuckGPT user script - https://greasyfork.org/en/scripts/459849-duckduckgpt

;------------------------------------------------------------------------------
;  = Clipboard Fn

;    + Clip_wait

Clip_wait(secs := 2, retrn := 0) {                                      ; Return = default False
ToolTipFn("Waiting for clipboard", secs * 1000)                         ; 2s
If not ClipWait(secs) {
    ToolTipFn(A_ThisHotkey ":: Clip_wait() failed!", 2000)              ; 2s
    ; MyNotification_show(A_ThisHotkey ":: Clip_wait() failed!", 2000)    ; 2s ; alternative to tooltip
    Exit
    }

If retrn = 1
    Return A_Clipboard
}

;--------
;    + Clip_call
; clipboard is emptied and new text is copied to clipboard

Clip_call(secs := 2, retrn := 0) {                                      ; Return = default False
ToolTipFn("Waiting for clipboard", secs * 1000)                         ; 2s
Global clipAll  := ClipboardAll()                                       ; Global = Return clipAll
A_Clipboard     := ""
Send "^c"
If not ClipWait(secs) {
    ToolTipFn(A_ThisHotkey ":: Clip_call() failed!", 2000)              ; 2s
    OnClipboardChange MultiClip_onChange, 0                             ; OFF, don't copy text to MultiClip_arr
    A_Clipboard := clipAll
    OnClipboardChange MultiClip_onChange, 1                             ; ON
    If retrn = 0
        Exit
    Else Return "err0r"                                                 ; If retrn = 1
    }
}

;--------
;    + Clip_callVar
; copied text is sent to variable, clipboard is restored

Clip_callVar(secs := 2, retrn := 0) {                                   ; Return = default False
OnClipboardChange MultiClip_onChange, 0                                 ; OFF, don't copy text to MultiClip_arr
If Clip_call(secs, retrn) == "err0r"                                    ; case-sensitive
    Return "err0r"                                                      ; MultiClip_onChange is turned on
Else {
    clipVar     := A_Clipboard
    A_Clipboard := clipAll
    OnClipboardChange MultiClip_onChange, 1                             ; ON
    Return clipVar
    }
}

;------------------------------------------------------------------------------
;  = ToolTip function

;    + ToolTipFn

ToolTipFn(mytext, myduration := 2000, xAxis?, yAxis?) { ; 2s
Try ToolTip()
; If WinExist("ahk_class tooltips_class32") ; alternative
;     WinClose()

ToolTip(mytext, xAxis?, yAxis?)
SetTimer () => ToolTip(), Abs(myduration) * -1 ; 1ms ; new thread ; always negative number
}

;------------------------------------------------------------------------------
;  = Wrap Text In Quotes or Symbols
; Inspired by https://geekdrop.com/content/super-handy-autohotkey-ahk-script-to-change-the-case-of-text-in-line-or-wrap-text-in-quotes
; and https://www.autohotkey.com/board/topic/9805-easy-encloseenquote/?p=61995

;    + WrapText_menuFn

WrapText_menuFn() {
If IsSet(WrapText_menu) {
    WrapText_menu.Show
    Return
    }
; Else ; Continue as below

Global WrapText_menu := Menu()      ; starts building a pop-up menu
; WrapText_menu.Delete()            ; replaced by If IsSet()

; add menu items
WrapText_menu.Add("&1  →  `'  Single Quotation `'"     , WrapText_selection )
WrapText_menu.Add("&2  →  `" Double Quotation `""      , WrapText_selection )
WrapText_menu.Add("&3  →  (  Round Brackets )"         , WrapText_selection )
WrapText_menu.Add("&4  →  [  Square Brackets ]"        , WrapText_selection )
WrapText_menu.Add("&5  →  {  Flower Brackets }"        , WrapText_selection )
WrapText_menu.Add("&6  →  ``  Accent/Backtick ``"      , WrapText_selection )
WrapText_menu.Add("&7  → `% Percent Sign `%"           , WrapText_selection )
WrapText_menu.Add("&8  →  ‘  Single Comma Quotation ’" , WrapText_selection )
WrapText_menu.Add("&9  →  “ Double Comma Quotation ”"  , WrapText_selection )
WrapText_menu.Add("&0  → Remove all"                   , WrapText_selection )

; add menu icons
MenuIcons_add(WrapText_menu, 10)

; show pop-up menu
WrapText_menu.Show
}

;--------
;    + WrapText_selection

WrapText_selection(ItemName, ItemPos, MyMenu) {
If ItemPos = 1
    WrapText_action( WrapText_vLead1 , WrapText_vTrail1 )     ; enclose in single quotation '' - ' U+0027 : APOSTROPHE
Else If ItemPos = 2
    WrapText_action( WrapText_vLead2 , WrapText_vTrail2 )     ; enclose in double quotation "" - " U+0022 : QUOTATION MARK
Else If ItemPos = 3
    WrapText_action( WrapText_vLead3 , WrapText_vTrail3 )     ; enclose in round brackets  ()
Else If ItemPos = 4
    WrapText_action( WrapText_vLead4 , WrapText_vTrail4 )     ; enclose in square brackets []
Else If ItemPos = 5
    WrapText_action( WrapText_vLead5 , WrapText_vTrail5 )     ; enclose in flower brackets {}
Else If ItemPos = 6
    WrapText_action( WrapText_vLead6 , WrapText_vTrail6 )     ; enclose in accent/backtick ``
Else If ItemPos = 7
    WrapText_action( WrapText_vLead7 , WrapText_vTrail7 )     ; enclose in percent sign %%
Else If ItemPos = 8
    WrapText_action( WrapText_vLead8 , WrapText_vTrail8 )     ; enclose in ‘’ - ‘ U+2018 LEFT & ’ U+2019 RIGHT SINGLE QUOTATION MARK {single turned comma & comma quotation mark}
Else If ItemPos = 9
    WrapText_action( WrapText_vLead9 , WrapText_vTrail9 )     ; enclose in “” - “ U+201C LEFT & ” U+201D RIGHT DOUBLE QUOTATION MARK {double turned comma & comma quotation mark}
Else                                                          ; ItemPos = 10
    WrapText_action( ""              , ""               )     ; remove above quotes
}

;--------
;    + WrapText_action

WrapText_action(q, p, direct := 0) { ; direct = default False

If direct = 1 && WinActive("ahk_class MSPaintApp") && ControlGetClassNN(ControlGetFocus()) != "RICHEDIT50W1" {
    ; enable wrapping text when WrapText_action is directly triggered in MSPaintApp when inserting/editing text element (ClassNN: RICHEDIT50W1)
    ; If not, then just Send the triggered hotkey
    Send A_ThisHotkey
    Exit
    }
; Else proceed with wrapping text

Clip_call(2)                                                    ; 2s, Exit
; use Clip_callVar instead of Clip_call earlier in the function to retain clipboard contents

TextString          := StrReplace(A_Clipboard, "`r")            ; remove \r for StrLen
TextStringInitial   := TextString                               ; save initial string for later
TextString          := RegExReplace(TextString,"^\s+|\s+$")     ; RegEx remove leading/trailing ; \s = [\r\n\t\f\v]
Len                 := StrLen(TextString)                       ; string length

; remove existing wrap characters, predefined Global variables for leading '"([{`%‘“ and trailing ”’%`}])"' characters
; Example: "Hello" → Hello     and  '"([{`%‘“Hello”’%`}])"' → Hello
; However, "Hello' → "Hello' ---- no change because leading and trailing characters don't match
Loop {
    If A_Index = 10
        Break

    ; only if wrap characters are the leading (position 1) and trailing characters (last position = length of string)
    ; Example: '"([{`%‘“Hel'lo”’%`}])"' → Hel'lo
    If InStr(TextString, WrapText_vLead%A_Index%) = 1 && InStr(TextString, WrapText_vTrail%A_Index%, , -1) = Len {
        TextString  := SubStr(TextString, 2, Len - 2)   ; remove wrap characters if found
        Len         := StrLen(TextString)               ; determine string length again and
        A_Index     := 0                                ; reset Loop for multiple wrapping characters if any
        }
    }
/* ; Alternative to Loop command
; this works even when string contains mixed leading and trailing wrap characters such as "Hello' → Hello
TextString := RegExReplace(TextString,'^[\[`'\(\{%`"“‘]+|^``',, &ReplacementCount)     ;"; remove leading  ['({%"“‘`  ; customise as your needs in WrapText_menuFn and WrapText Keys
If ReplacementCount > 0 ; don't remove trailing character if leading character is not removed
    TextString := RegExReplace(TextString,'[\]`'\)\}%`"”’]+$|``$')     ;"; remove trailing ]')}%"”’`  ; customise as your needs in WrapText_menuFn and WrapText Keys
*/

; add new wrapping characters and determine length again
TextString  := q TextString p
Len         := StrLen(TextString)

; If you regularly include leading/trailing spaces within quotes, comment out above RegEx and below If statements
If RegExMatch(TextStringInitial, "^\s+", &Lead) {   ; If the initial string has leading \s
    TextString := Lead[] TextString                 ; add &OutputVar to string
    Len        += Lead.Len                          ; add the length of &OutputVar to Len
    }
If RegExMatch(TextStringInitial, "\s+$", &Trail) {  ; If the initial string has trailing \s
    TextString .= Trail[]                           ; append &OutputVar to string
    Len        += Trail.Len                         ; add the length of &OutputVar to Len
    }

; Send "{Raw}" TextString                           ; send the string with quotes
pasteThis(TextString)                               ; paste
If Len <= Send_LenLimit                             ; and select text string if ≤ Send_LenLimit
    Send "+{Left " Len "}"                          ; change `Send_LenLimit` as needed, bigger Len = longer execution time
; A_Clipboard := TextStringInitial                    ; restore original text string to clipboard if desired
}

;------------------------------------------------------------------------------
;  = Base64 Encode/Decode
; 2024.02.03 modified from https://www.autohotkey.com/boards/viewtopic.php?p=555827#p555827

;    + Base64_encode

Base64_encode(String := A_Clipboard) {

; Convert the input string into a byte string of UTF-8 characters.
size    := StrPut(String, "UTF-8")
bin     := Buffer(size)
StrPut(String, bin, "UTF-8")
size    := size - 1                     ; A binary does not have a null terminator

; Calculate the length of the base64 string.
length  := 4 * Ceil(size / 3) + 1       ; A string has a null terminator
VarSetStrCapacity(&str, length)         ; Allocates a ANSI or Unicode string
; This appends 1 or 2 zero byte null terminators respectively.

; Passing a pre-allocated string buffer prevents an additional memory copy via StrGet.
flags   := 0x40000001                   ; CRYPT_STRING_NOCRLF | CRYPT_STRING_BASE64
; If not (DllCall("crypt32\CryptBinaryToString", "ptr", bin, "uint", size, "uint", flags, "str", str, "uint*", &length))
;     ToolTipFn("Error: B64 Encoding Failed", 2000) ; 2s

; error text if needed
info    := "B64 encoding failed!" "`n" "String: " String "`n"

Try {
    If not DllCall("crypt32\CryptBinaryToString", "ptr", bin, "uint", size, "uint", flags, "str", str, "uint*", &length) {
        SetTimer () => MsgBox(info,, 262144), -1                            ; 1ms ; new thread ; 262144 Always-on-top
        Return
        }
    }
Catch as err
    SetTimer () => CatchError_show( CatchError_details(err, info) ), -1     ; 1ms ; new thread
Else Return str
}

;--------
;    + Base64_decode

Base64_decode(Base64 := A_Clipboard) {

; Trim whitespace and remove mime type.
string       := RegExReplace(Trim(Base64), "(?i)^.*?;base64,")

; Retrieve the size of bytes from the length of the base64 string.
size    := StrLen(RTrim(string, "=")) * 3 // 4
bin     := Buffer(size)

; error text if needed
info    := "B64 decoding failed!" "`n" "B64 String: " Base64 "`n"

; Place the decoded base64 string into a binary buffer.
flags   := 0x1 ; CRYPT_STRING_BASE64
Try {
    If not DllCall("crypt32\CryptStringToBinary", "str", string, "uint", 0, "uint", flags, "ptr", bin, "uint*", size, "ptr", 0, "ptr", 0) {
        SetTimer () => MsgBox(info,, 262144), -1                            ; 1ms ; new thread ; 262144 Always-on-top
        Return
        }
    }
Catch as err
    SetTimer () => CatchError_show( CatchError_details(err, info) ), -1     ; 1ms ; new thread
Else Return StrGet(bin, size, "UTF-8")                                      ; Must reinterpret the binary bytes from UTF-8
}

;--------
;    + Base64_decodeX2

Base64_decodeX2(b64X2 := A_Clipboard) {
str1 := Base64_decode(b64X2)
If str1 !== b64X2 {                     ; case-sensitive
    str2 := Base64_decode(str1)
    If str2 !== str1                    ; case-sensitive
        Return Trim(str1 "`n" str2, "`r`n`t`s`v`f")
    }
Else Return str1
}

;------------------------------------------------------------------------------
;  = URL Encode/Decode
; Modified from https://www.autohotkey.com/boards/viewtopic.php?style=7&t=116056#p517193
; decodedURL RegEx modified from https://www.makeuseof.com/regular-expressions-validate-url/

;    + URL_decode

URL_decode(Url, Enc := "UTF-8") {

originUrl   := Url
decodedURL  := "^http(s|)://[\w-@:%~#&=\.\+\?]{2,256}\.[a-z]{2,6}($|/[\w-@:%~#&=/\.\+\?]*$)"

Pos         := 1
Loop {
    Pos     := RegExMatch(Url, "i)(?:%[\da-f]{2})+", &code, Pos++) ; %3A
    If Pos = 0
        Break
    var     := Buffer(StrLen(code[0]) // 3, 0)
    code    := SubStr(code[0], 2)
    Loop Parse code, "`%"
        NumPut("UChar", Integer("0x" A_LoopField), var, A_Index - 1)
    Url     := StrReplace(Url, "`%" code, StrGet(var, Enc))
    }

If originUrl != Url {                                   ; input ≠ output
    ToolTipFn("URL_decode successful!", 2000)           ; 2s
    Return Url
    }
Else If not RegExMatch(Url, decodedURL)                 ; If input = output and not a RegEx match to decoded url
    MsgBox( "URL_decode failed!" "`n`n"
        .   "String: " originUrl "`n`n"
        .   "Not an encoded/decoded URL"    ,, 262144)  ; 262144 Always-on-top
Else ToolTipFn("URL already decoded!", 2000)            ; 2s
Exit                                                    ; since no change is made to url, stop further action
}

;--------
;    + URL_encode

URL_encode(str, sExcepts := "-_.", enc := "UTF-8") {

encodedURL := "^http(s|)%3A%2F%2F[\w-@%~#&=\.\+\?]{2,256}\.[a-z]{2,6}($|%2F[\w-@%~#&=\.\+\?]*$)"
decodedURL := "^http(s|)://[\w-@:%~#&=\.\+\?]{2,256}\.[a-z]{2,6}($|/[\w-@:%~#&=/\.\+\?]*$)"

; Validate input
If not RegExMatch(str, decodedURL) {
    ; input ≠ decoded url, check If already encoded
    If not RegExMatch(str, encodedURL)
        MsgBox( "URL_encode failed!"    "`n`n"
            .   "String: " str          "`n`n"
            .   "Not an encoded/decoded URL"    ,, 262144)  ; 262144 Always-on-top
    Else
        ToolTipFn("URL already encoded!", 2000)             ; 2s

    Exit
    }
; Else ; input = decoded url, proceed to encode

hex     := "00"
buff    := Buffer(StrPut(str, enc))     ; ReqBufSize
StrPut(str, buff, enc)                  ; copy to buffer
encoded := ""
Loop {
    If not b := NumGet(buff, A_Index - 1, "UChar")
        Break
    ch := Chr(b)
    If  (  b >= 0x41 && b <= 0x5A       ; A-Z
        || b >= 0x61 && b <= 0x7A       ; a-z
        || b >= 0x30 && b <= 0x39       ; 0-9
        || InStr(sExcepts, ch, 1)  )    ; CaseSense = True
        encoded .= ch
    Else {
        DllCall("msvcrt\swprintf", "Str", hex, "Str", "%%%02X", "UChar", b, "Cdecl")
        encoded .= hex
        }
    }
ToolTipFn("URL_encode successful!", 2000) ; 2s
Return encoded
}

;------------------------------------------------------------------------------
;  = MouseMove Fn

;    + MouseMove_screenCenter

MouseMove_screenCenter() {
CoordMode "Mouse", "Screen"                     ; Screen is default for CoordMode command parameter
MouseMove A_ScreenWidth / 2, A_ScreenHeight / 2
CoordMode "Mouse", "Client"                     ; Client is default for most AHK commands
}

;--------
;    + MouseMove_clientCenter

MouseMove_clientCenter(HWND := "A") {
WinGetClientPos , , &OutWidth, &OutHeight, HWND
MouseMove OutWidth / 2, OutHeight / 2           ; Coordinates are relative to the active window's client area
}

;------------------------------------------------------------------------------
;  = Kill All Instances Of An App

;    + KillAll_titles

KillAll_titles(HWNDarr) {
Titles := ""
Loop HWNDarr.Length {
    t := RegExReplace(WinGetTitle(HWNDarr[A_Index]), "\s+", A_Space)   ; remove multiple \s = [\r\n\t\f\v]

    If t ~= "^ *$" ; RegEx match to empty or only Spaces
        Titles .= Format("{:5}", A_Index) " = #No Title"      "`n"
    Else Titles .= Format("{:5}", A_Index) " = " t            "`n"
    }
Return Titles
}

;------------------------------------------------------------------------------
;  = Print Screen Fn

;    + ScreenSnip_menuFn

ScreenSnip_menuFn() {
If IsSet(ScreenSnip_menu) {
    ScreenSnip_menu.Show
    Return
    }
; Else ; Continue as below

Global ScreenSnip_menu := Menu()   ; starts building a pop-up menu
; ScreenSnip_menu.Delete()         ; replaced by If IsSet()

; add menu items
ScreenSnip_menu.Add("&1  → Rectangular Snip"          , ScreenSnip_selection )
ScreenSnip_menu.Add("&2  → Window Snip"               , ScreenSnip_selection )
ScreenSnip_menu.Add("&3  → Full Screen Snip"          , ScreenSnip_selection )
ScreenSnip_menu.Add("&4  → Freeform Snip"             , ScreenSnip_selection )

; add menu icons
MenuIcons_add(ScreenSnip_menu, 4)

; show pop-up menu
ScreenSnip_menu.Show
}

;--------
;    + ScreenSnip_selection
; Applicable to SnippingTool.exe version 11.2407.3.0 and later (Date: 2024.09.16)

ScreenSnip_selection(ItemName, ItemPos, MyMenu) {
If ItemPos = 3 {
    ScreenSnip_printScreen()
    Return
    }

Send "{PrintScreen}"
If WinWaitActive("Snipping Tool Overlay ahk_exe SnippingTool.exe",, 3) { ; 3s
    Sleep 250
    MouseMove A_ScreenWidth / 2, 40 ; client 973, 38
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
If WinWaitClose("Snipping Tool Overlay ahk_exe SnippingTool.exe",, 15) && (A_PriorKey != "Escape") { ; 15s ; case insensitive

    ; save time to variable because even a 1 sec delay in formatting can cause error
    sTime := A_Now

    ; If `Automatically save screenshots` is ENABLED in snippingtool
    SetTimer () => ScreenSnip_fileOp(FormatTime(sTime, "yyyy-MM-dd HHmmss")), -1   ; 1ms ; new thread

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

Else If A_PriorKey = "Escape"
    ToolTipFn(A_ThisHotkey ":: Screen Snipping aborted - Esc", 2000)            ; 2s
Else ToolTipFn(A_ThisHotkey ":: Screen Snipping aborted - 15s timeout", 2000)   ; 2s
}

/* For older versions of Snipping tool, below code may work

ScreenSnip_selection(ItemName, ItemPos, MyMenu) {
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
;    + ScreenSnip_printScreen

ScreenSnip_printScreen() {

; save screenshot number in variable 'serial'
serial := RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer", "ScreenshotIndex", "1")

; send Windows + Print Screen
Send "#{PrintScreen}"

; use SetTimer to wait for file creation 100ms and execute action in another thread
; pass screenshot number obtained from registry as number for path of screenshot file in `ScreenSnip_fileOp` function
SetTimer () => ScreenSnip_fileOp("(" serial ")"), -1     ; 1ms ; new thread
}

;--------
;    + ScreenSnip_fileOp

ScreenSnip_fileOp(id) {

; generate path where screenshot is likely to be saved
savedPath := "C:\Users\" A_UserName "\Pictures\Screenshots\Screenshot " id ".png"

; if file does not exist or not completely saved, wait using Sleep and repeat check
While not (FileExist(savedPath) && FileGetSize(savedPath) > 0) {
    Sleep 500                       ; wait for 500ms
    If A_Index > 6 {                ; max 6 attempts, total wait 3000ms = 3s

        ; If failed ×6, open screenshot folder in explorer
        OpenFolder("C:\Users\" A_UserName "\Pictures\Screenshots")

        ; show generated path in pop-up to compare with actual file path
        ShowErrExit(A_ThisHotkey ":: Screenshot not saved or zero size after 3s!" "`n"
                .   savedPath                                                           ) ; give up on rename
        }
    }

; prepare to rename file
fileTime := FormatTime(FileGetTime(savedPath, "C"), "yyyy-MM-dd @ HH：mm：ss") ; 2016-07-21 @ 13：28：05
    ; FileGetTime - obtain creation time as a string in YYYYMMDDHH24MISS format
    ; FormatTime - transform Timestamp YYYYMMDDHH24MISS into desired date/time format.

NewPath := "C:\Users\" A_UserName "\Pictures\Screenshots\" fileTime ".png"

FileMove savedPath, NewPath                                             ; rename

; Further actions - (uncomment below lines to execute)
; Run 'mspaint.exe "' NewPath '"',,"Max"                                  ; open in paint
; If WinWaitActive("ahk_class MSPaintApp",, 3) {                          ; 3s ; wait for paint window to open
;     Sleep 500                                                           ; 500ms ; wait for MSPaintApp to load
;     Send "!3"           ; activate 3rd tool in Quick Access Toolbar     ; "Select" tool as personal preference
;     }
; OpenFolder("C:\Users\" A_UserName "\Pictures\Screenshots\")             ; open screenshot folder in explorer
}

;------------------------------------------------------------------------------
;  = Toggle Window Always-On-Top Fn

;    + AlwaysOnTop_enable

AlwaysOnTop_enable(HWND := WinGetID("A"), t := WinGetTitle(HWND), onTopTitle := "! ") {
Try {
    WinSetAlwaysOnTop 1, HWND           ; Turn ON always on top
    If t != "" && onTopTitle != SubStr(t, 1, StrLen(onTopTitle))
        WinSetTitle onTopTitle t, HWND  ; add onTopTitle to window title
    }
Catch as err
    SetTimer () => CatchError_show( CatchError_details(err, "AlwaysOnTop_enable() failed!") ), -1   ; 1ms ; new thread
}

;--------
;    + AlwaysOnTop_disable

AlwaysOnTop_disable(HWND := WinGetID("A"), t := WinGetTitle(HWND), onTopTitle := "! ") {
Try {
    WinSetAlwaysOnTop 0, HWND       ; Turn OFF always on top
    If t != "" && onTopTitle = SubStr(t, 1, StrLen(onTopTitle))
        ; change title only if title is not empty, because some windows don't have a title/title bar
        ; i.e., don't run WinSetTitle for windows such as 'ahk_class #32770 ahk_exe SndVol.exe'
        ; remove onTopTitle from window title
        WinSetTitle StrReplace(t, onTopTitle), HWND
    }
Catch as err
    SetTimer () => CatchError_show( CatchError_details(err, "AlwaysOnTop_disable() failed!") ), -1   ; 1ms ; new thread
}

;------------------------------------------------------------------------------
;  = Text correction by Menu
; inspired by https://jacksAutoHotkeyblog.wordpress.com/2019/11/04/AutoHotkey-hotstring-menus-for-text-replacement-options/#more-40700
; and https://jacksAutoHotkeyblog.wordpress.com/2015/10/22/how-to-turn-AutoHotkey-hotstring-autocorrect-pop-up-menus-into-a-function-part-5-beginning-hotstrings/

;    + TextCorrection_menuFn

TextCorrection_menuFn(TextOptions) {

opt := StrSplit(TextOptions, ",", " `t")
; generate array from string, omit Spaces and Tabs from the beginning and end (but not the middle) of every item

Global TextCorrection_menu := Menu()        ; starts building a pop-up menu
TextCorrection_menu.Delete()                ; retained because TextOptions may change

; add menu items
Loop opt.Length
    TextCorrection_menu.Add("&" MenuShortcuts[A_Index] "  → " opt[A_Index], TextCorrection_action)

; add menu icons
MenuIcons_add(TextCorrection_menu, opt.Length)

; show pop-up menu
TextCorrection_menu.Show

;--------
;    + TextCorrection_action (enclosed Fn)

    TextCorrection_action(ItemName, ItemPos, MyMenu) {
    pasteThis( opt[ItemPos] )
    }
}

;------------------------------------------------------------------------------
;  = File operations

;    + Folder_move

Folder_move(sourceDir, destDir, overwrite := 0, nameExclude := "^.") {    ; overwrite files = default False

errors      := ""   ; store errors from Loop
dirCount    := 0    ; start at zero

Loop Files sourceDir "\*", "D" {
    If not RegExMatch(A_LoopFileName, nameExclude) {
        Try DirMove A_LoopFileFullPath, destDir "\" A_LoopFileName , overwrite
        Catch as err
            errors .= CatchError_details(err, A_Index ": " A_LoopFileFullPath) "`n"
        Else dirCount++
        }
    }
If errors != "" {
    OpenFolder(sourceDir)
    Global errorDetails .= 'Folder_move("' sourceDir '", "' destDir '", "' overwrite '", "' nameExclude '")' "`n" errors "`n"
    }
Return dirCount
}

;------------------------------------------------------------------------------
;  = String Fn

;    + CleanText_trim
; remove leading/trailing Spaces, `r, double Spaces and trim

CleanText_trim(myText) {
trimTxt := StrReplace(  myText  , "`r"          )   ; trim CR
trimTxt := RegExReplace(trimTxt , "m)^ +| +$"   )   ; multi-line, trim leading/trailing spaces
trimTxt := RegExReplace(trimTxt , "  +", A_Space)   ; replace multiple Spaces with single Space
Return     Trim(        trimTxt , "`r`n`t`s`v`f")   ; trim leading/trailing linefeed, Tab, space, carriage return, vertical tab, formfeed
}

;--------
;    + Str_LineLimit
; limit lines to 35

Str_LineLimit(string, LineLimit := 35) {

str     := Trim(string, "`r`n`t`s`v`f") ; trim input
sArr    := StrSplit(str, "`n", "`r")

If sArr.Length > LineLimit {
    output := ""
    Loop (LineLimit - 1)
        output .= sArr[A_Index] "`n"
    Return output "… and more!"
    }
Else Return str
}

;--------
;    + Str_LenLineimit
; limit each line length to 150 characters (because Tooltip's limit is ~150 without being too large)
; and limit total number of lines to 35 (because MsgBox's limit is ~43 with display resolution of 1080p)

Str_LenLineimit(string, LenLimit := 150, LineLimit := 35) {

str     := Trim(string, "`r`n`t`s`v`f") ; trim input
sArr    := StrSplit(str, "`n", "`r")

; trim number of lines
If sArr.Length > LineLimit {
    sArr.Length := LineLimit - 1
    sArr.Push("… and more!")
    }

; trim line length
output := ""
Loop sArr.Length {
    If StrLen(sArr[A_Index]) > LenLimit
        output .= SubStr(sArr[A_Index], 1, LenLimit - 2) " …" "`n"
    Else output .= sArr[A_Index] "`n"
    }

Return Trim(output, "`r`n`t`s`v`f")
}

;------------------------------------------------------------------------------
;  = Windows Explorer Fn

;    + Explorer_GetSize

Explorer_GetSize(pathType, pathContent) {

; variables
SizeB           := 0
errorDetails    := ""
folderName      := ""
fileName        := ""
files           := ""
fileCount       := 0

If pathType = 1 {                                                               ; multiple folder/files
    Loop Parse pathContent, "`n", "`r" {
        If DirExist(A_LoopField) {                                              ; is folder
            Loop Files A_LoopField "\*", "R" {
                SizeB += A_LoopFileSize
                fileCount++
                }
            folderName .= A_LoopField "`n"
            }
        Else If FileExist(A_LoopField) {                                        ; is file
            SizeB += FileGetSize(A_LoopField, "B")
            fileName .= A_LoopField "`n"
            fileCount++
            }
        Else errorDetails .= "#" A_Index ": " A_LoopField " → Not a folder or file" "`n"
        }

    ; display folders first in pathContent
    If fileName = ""                                                            ; no files, only folders
        pathContent := Trim(Sort(folderName, "CLogical"), "`r`n`t`s`v`f")
    Else If folderName = ""                                                     ; no folders, only files
        pathContent := Trim(Sort(fileName  , "CLogical"), "`r`n`t`s`v`f")
    Else                                                                        ; both files and folders
        pathContent := Trim(Sort(folderName, "CLogical"), "`r`n`t`s`v`f") "`n"  ; folders +
                    .  Trim(Sort(fileName  , "CLogical"), "`r`n`t`s`v`f")       ; files
    }

Else If pathType = 2 {                                                          ; single folder
    Loop Files pathContent "\*", "R" {
        SizeB += A_LoopFileSize
        files .= A_LoopFileFullPath " (" SizeFn(A_LoopFileSize) ")" "`n"
        fileCount++
        }
    }

Else If pathType = 3 {                                                          ; single file
    SizeB += FileGetSize(pathContent, "B")
    fileCount++
    }
Else ; pathType = (4) CLSID or (0) invalid mono/multi-line
    errorDetails .= "Explorer_GetSize() error! Incorrect pathType: " pathType

Return [SizeB, errorDetails, pathContent, Trim(Sort(files, "CLogical"), "`r`n`t`s`v`f"), fileCount]
}

;--------
;    + RBinQuery
; Modified from https://www.autohotkey.com/boards/viewtopic.php?p=125495#p125495

RBinQuery(userDriveInput := "") {
Global errorDetails := ""                                       ; store errors
RBinArr             := []                                       ; array to store return
driveLetters        := checkDrivesFn(userDriveInput)

If userDriveInput = ""
    userDriveInfo := "User Query - Blank"
Else userDriveInfo := 'User Query - "' userDriveInput '"'

If userDriveInput = "" && errorDetails = ""
    ErrorReturn := ""
Else If errorDetails = ""
    ErrorReturn := "No errors in user query!"       "`n"
Else ErrorReturn := "Errors in user query:"         "`n" errorDetails

If driveLetters = ""
    ShowErrExit(A_ThisHotkey ":: RBinQuery() No matches to drive letters found!"    "`n`n"
            .   userDriveInfo                                                       "`n"
            .   ErrorReturn                                                                 )

SID := userSID()

; list of folders/files visible in windows file explorer
visibleArr := RBinVisible(driveLetters)                         ; returns [visibleSummary, visibleRBFCount, visibleTotalSizeB]

; list of folders/files *NOT* visible in windows file explorer
hiddenArr := RBinHidden(driveLetters, SID, visibleArr[1])       ; returns [RBSummary, TotalRBFCount, TotalSizeB, hiddenFileList]

RBinArr.InsertAt(1 , driveLetters                           )
RBinArr.InsertAt(2 , visibleArr[1]                          )   ; visibleSummary
RBinArr.InsertAt(3 , visibleArr[2]                          )   ; visibleRBFCount   (used for hiddenRBFCount)
RBinArr.InsertAt(4 , SizeFn( visibleArr[3] )                )   ; visibleTotalSizeB (used for hiddenTotalSizeB)

RBinArr.InsertAt(5 , hiddenArr[1]                           )   ; RBSummary
RBinArr.InsertAt(6 , hiddenArr[2]                           )   ; TotalRBFCount
RBinArr.InsertAt(7 , SizeFn( hiddenArr[3] )                 )   ; TotalSizeB

RBinArr.InsertAt(8 , hiddenArr[2] - visibleArr[2]           )   ; hiddenRBFCount
RBinArr.InsertAt(9 , SizeFn( hiddenArr[3] - visibleArr[3] ) )   ; hiddenTotalSizeB
RBinArr.InsertAt(10, hiddenArr[4]                           )   ; hiddenFileList

RBinArr.InsertAt(11, userDriveInfo                          )   ; userDriveInfo
RBinArr.InsertAt(12, ErrorReturn                            )   ; ErrorReturn

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
visibleSummary := "; List of folders/files visible in File Explorer Recycle Bin -"              "`n`n"
               .  "Number of visible items     : " visibleRBFCount                              "`n"
               .  "Total size of visible items : " SizeFn(visibleTotalSizeB)                    "`n`n"

For item in visibleFiles {
    visibleSummary .= "#" Format("{:03d}", A_Index) " Recycle Bin Location   : " item.Location  "`n"
                   .  "     Original Location      : " item.OriginalLocation "\" item.Name      "`n"
                   .  "     File Size              : " SizeFn(item.FileSize)                    "`n`n"
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
    Loop Files DriveL ":\$Recycle.Bin\" SID "\*", "R" {
        RBFCount++
        SizeB += A_LoopFileSize
        If not InStr(visibleSummary, A_LoopFileFullPath) ; exclude file name if already in visibleSummary
            FileNames .= Format("{:-80}", A_LoopFileFullPath) "(" SizeFn(A_LoopFileSize) ")" "`n"
        }
    RBSummary .= "Drive: " DriveL               "`n"
              .  "No. of files: " RBFCount      "`n"
              .  "Size: " SizeFn(SizeB)         "`n`n"
    TotalRBFCount   += RBFCount
    TotalSizeB      += SizeB
    }

hiddenFileList := Trim(Sort(FileNames, "CLogical"), "`r`n`t`s`v`f")
Return [RBSummary, TotalRBFCount, TotalSizeB, hiddenFileList]
}

;--------
;    + userSID

userSID() {
    If IsSet(SID)
        Return SID
    ; Else ; proceed

    Loop Reg "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList", "R KV" {
        If (A_LoopRegName == "ProfileImagePath") {  ; case-sensitive
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
Global errorDetails            ; don't add := because it should be assigned by calling function/hotkey/hotstring
driveList := DriveGetList()

If userDriveInput = ""         ; If drive letters are not specified or not valid
    Return driveList           ; get list of drives - e.g., ACDEZ
Else {
    checkedDrives := ""
    Loop Parse userDriveInput {
        If InStr(driveList, A_LoopField) ; use InStr because RegEx match (~=) is case sensitive
            checkedDrives .= A_LoopField "`n"
        Else errorDetails .= ' * Bad Match - "' A_LoopField '" → Available drives: ' driveList "`n"
        }
    Return StrReplace(Sort(checkedDrives, "CLogical U"), "`n") ; Case-insensitive sort, remove U duplicates and `n line feed
    }
}

;--------
;    + RBinDisplay

RBinDisplay(RBinArr) {

msgTitle := "RBinQuery: " RBinArr[1]                                                ; driveLetters
msgText := RBinArr[11]                                              "`n"            ; userDriveInfo
        .  RBinArr[12]                                              "`n"            ; ErrorReturn
        .  "# Below summary includes all files (visible + hidden)"  "`n`n"
        .  RBinArr[5] "Total No. of files: " RBinArr[6]             "`n"
        .  "Size: " RBinArr[7]                                                      ; RBSummary, totals

Result1 := MsgBox_Custom(msgText, msgTitle, , 2, "File List", "OK")                 ; 2 = Yes / No

If Result1 = "Yes" {                                                                ; renamed button - File List
    rbFile  := A_Temp "\List of files in Recycle bin.txt"
    Results := RBinGenerateTxt(msgTitle, msgText, RBinArr)
    FileCreate_Overwrite(rbFile, Results)
    Run '"' rbFile '"',,"Max" ;"
    }
}

;--------
;    + RBinGenerateTxt

RBinGenerateTxt(msgTitle, msgText, RBinArr) {
Results := " ;"; continuation section
(
; * Table of contents *
; Summary
; List of folders/files visible in File Explorer Recycle Bin -
; List of hidden folders/files in Recycle Bin -

;------------------------------------------------------------------------------
; Summary

)" . msgTitle                                                                           "`n`n"  ;"; 1 driveLetters
   . msgText                                                                            "`n`n"  ; 5 RBSummary, 6 TotalRBFCount, 7 TotalSizeB
   . ";------------------------------------------------------------------------------"  "`n"
   . RBinArr[2]                                                                                 ; visibleSummary
   . ";------------------------------------------------------------------------------"  "`n"
   . "; List of hidden folders/files in Recycle Bin -"                                  "`n`n"
   . "Number of hidden items     : " RBinArr[8]                                         "`n"    ; hiddenRBFCount
   . "Total size of hidden items : " RBinArr[9]                                         "`n`n"  ; hiddenTotalSizeB
   . RBinArr[10]                                                                                ; hiddenFileList

Return Results
}

;--------
;    + SizeFn

SizeFn(SizeB) {
If SizeB >= 1024 {
    readableSizeKB := SizeB / 1024
    If readableSizeKB >= 1024 {
        readableSizeMB := readableSizeKB / 1024
        If readableSizeMB >= 1024 {
            readableSizeGB := readableSizeMB / 1024
            If readableSizeGB >= 1024
                Return Round(readableSizeGB / 1024, 2) " TB"
            Else Return Round(readableSizeGB, 2) " GB"
            }
        Else Return Round(readableSizeMB, 2) " MB"
        }
    Else Return Round(readableSizeKB, 2) " KB"
    }
Else Return SizeB " bytes"
}

;--------
;    + DeleteEmptyFolder

DeleteEmptyFolder(pathContent, pathType?) {

If IsSet(pathType)                          ; path already validated
    If pathType != 2                        ; path not (2) single folder, Return
        Return "DeleteEmptyFolder() aborted! pathType is not a single folder!" "`n" pathContent "`n`n"
    ; Else pathType = 2                     ; proceed
Else                                        ; path not validated
    If not DirExist(pathContent)            ; validation failed, Return
        Return "DeleteEmptyFolder() aborted! Path is not valid!" "`n" pathContent "`n`n"
    ; Else Dir valid                        ; proceed

result := Explorer_GetSize(2, pathContent)  ; pathType = 2 ; Returns [SizeB, errorDetails, pathContent, files, fileCount]

If result[1] = 0 {                          ; If SizeB = zero
    Try FileRecycle pathContent             ; delete empty folder, recursive
    Catch as err
        Return CatchError_details(err, "DeleteEmptyFolder() failed:" "`n" pathContent "(0 bytes)")  "`n`n"
    Return "DeleteEmptyFolder() success:" "`n"    pathContent                                       "`n`n"
    }
Else { ; Return folder with SizeB and files
    info := "DeleteEmptyFolder() aborted!"              "`n"
         .  pathContent " (" SizeFn(result[1]) ")"      "`n"    ; SizeB
         .  result[5] " files"                          "`n"    ; fileCount
         .  "Contents:"                                 "`n"
         .  result[4]                                   "`n`n"  ; files
    Return info
    }
}

;--------
;    + Explorer_ExtractFolder

Explorer_ExtractFolder(SourceFolder) {
; example for SourceFolder "D:\Movies\"

SourceFolder := RegExReplace(SourceFolder, "\\$") ; remove trailing "\", Result → D:\Movies

If not DirExist(SourceFolder) ; revalidation of path after RegExReplace
    Return "Not a valid SourceFolder: " SourceFolder "`n"
; Else Dir is valid ; proceed

; remove folder "Movies" and set "D:\" as the destination
DestinationFolder := RegExReplace(SourceFolder, "\\[^\\]+$")    ; Result → D:\

If not DirExist(DestinationFolder)
    Return "Not a valid DestinationFolder!" "`n" "Source: " SourceFolder "`n" "Destination: " DestinationFolder "`n"
; Else Dir is valid ; proceed

extraction_errors := ""

; move files
Try FileMove SourceFolder "\*", DestinationFolder
Catch as err
    extraction_errors .= CatchError_details(err, "Source - " SourceFolder "`n" "Destination: " DestinationFolder) "`n"

Global errorDetails := ""

; move folders
Folder_move(SourceFolder, DestinationFolder)

; add DirMove errors to extraction errors
extraction_errors   .= errorDetails
Global errorDetails := ""

If extraction_errors != "" {
    result  := Explorer_GetSize(2, SourceFolder) ; Returns [SizeB, errorDetails, pathContent, files, fileCount]

    info    := "SourceFolder not deleted! Remaining folder size: " SizeFn(result[1])    "`n`n"   ; SizeB
            .  "Extraction errors:"                                                     "`n"
            .  extraction_errors                                                        "`n"
            .  result[5] " files"                                                       "`n"     ; fileCount
            .  "Contents:"                                                              "`n"
            .  result[4]                                                                "`n"     ; files
    If result[2]                                                                                 ; errorDetails
        info .= "Explorer_GetSize() errors:"                                            "`n"
             .  result[2]                                                                        ; errorDetails
    Return info
    }
Else ; delete empty folder after extraction of all contents ; pathType = 2 single folder
    Return DeleteEmptyFolder(SourceFolder, 2) ; Returns aborted, success, error info
}

;--------
;    + OpenFolder

OpenFolder(newPath) {
If DirExist(newPath) {                                              ; if valid directory/folder
    If hwnd := WinExist("File Explorer ahk_class CabinetWClass") {  ; if explorer is open
        WinActivate
        ExplorerTab_Navigate(newPath, hwnd)                         ; navigate to newPath on topmost explorer
        }
    Else Run 'explore "' newPath '"',,"Max"
    }
Else If RegExMatch(newPath, "\{\w{8}-\w{4}-\w{4}-\w{4}-\w{12}\}", &Match) {
    ;  CLSID like {21EC2020-3AEA-1069-A2DD-08002B30309D}  All Control Panel Items

    Try Run "::" match[0]
    Catch {
        Details := A_ThisHotkey ":: OpenFolder() switched because of CLSID match"   "`n"
                .  "Path: "        newPath                                          "`n"
                .  "RegExMatch: "  match[0]                                         "`n"
                .  "1. Run(::) failed - OSError" OSError(A_LastError).Message       "`n"

        Loop_runShell()
        }

    /* alternative for CLSID that open within explorer - doesn't always work (e.g., Windows Tools CLSID {D20EA4E1-3957-11D2-A40B-0C5020524153})
    If WinExist("File Explorer ahk_class CabinetWClass") && Explorer_FocusAddressBar() {  ; Return = 1
        pasteThis("::" newPath)
        WinWaitClose("PopupHost ahk_class Microsoft.UI.Content.PopupWindowSiteBridge",, 2)
        ; 2s - wait for drop down to disappear, then Send Enter
        ; WinWait commands used to prevent drop down display appearing after Enter - explorer bug
        Send "{Enter}"                    ; focus file list
        ControlFocus "DirectUIHWND2", "File Explorer ahk_class CabinetWClass"
        }
    */

    }
Else ShowErrExit(A_ThisHotkey ":: OpenFolder() failed - newPath is not valid!"          "`n"
            .    newPath                                                                     )

;--------
;    + Loop_runShell (enclosed Fn)

    Loop_runShell() {
    success := 0
    shell   := "shell:"
    While success = 0 {
        Try Run shell match[0]
        Catch {
            Details .= A_Index + 1 ". Run(" shell ") failed - OSError" OSError(A_LastError).Message   "`n"
            shell   .= ":"
            }
        Else success := 1
        If A_Index > 2
            Break
        }
    If success = 0
        ShowErrExit(Details)
    Else Return Details
    }
}

;--------
;    + ExplorerTab_GetIdentity
; modified from https://old.reddit.com/r/AutoHotkey/comments/10fmk4h/get_path_of_active_explorer_tab/kuplyts/

ExplorerTab_GetIdentity(hwnd := WinExist("A")) {

win_class   := WinGetClass(hwnd)
Explorers   := ComObject("Shell.Application").Windows

If win_class = "Progman"                                ; case-INsensitive for entire function
    Return Explorers.Item(ComValue(0x13, 0x8))          ; 0x13 = VT_UI4 32-bit unsigned int, 0x8 = SWC_DESKTOP

Else If win_class != "CabinetWClass" && hwnd != WinExist("ahk_class CabinetWClass")
    ShowErrExit(A_ThisHotkey ":: Desktop / File Explorer is not the active window! ExplorerTab_GetIdentity() aborted!"  "`n"
            .   "Title: "       WinGetTitle(hwnd)                                                                       "`n"
            .   "ahk_class: "   win_class                                                                               "`n"
            .   "ahk_exe: "     WinGetProcessName(hwnd)                                                                     )

Else ; If win_class != "CabinetWClass" but window exists
    WinActivate(hwnd)
; Else win_class = "CabinetWClass", proceed as below

ClassNN     := ControlGetClassNN(ControlGetFocus(hwnd))

If ClassNN = "SysTreeView321" {
    mytext := "ExplorerTab_GetIdentity()" "`n" "ControlFocus changed!" "`n" "Navigation Pane → File List"
    SetTimer () => MyNotification_show(mytext, 5000,,, 1), -1   ; 1ms ; new thread ; notification = 5s
    ControlFocus "DirectUIHWND2", hwnd
    ClassNN := ControlGetClassNN(ControlGetFocus(hwnd))
    }

If ClassNN = "Microsoft.UI.Content.DesktopChildSiteBridge1" {
    mytext := "ExplorerTab_GetIdentity()" "`n" "ControlFocus changed!" "`n" "Command Bar → File List"
    SetTimer () => MyNotification_show(mytext, 5000,,, 1), -1   ; 1ms ; new thread ; notification = 5s
    ControlFocus "DirectUIHWND2", hwnd
    Sleep 500 ; 500ms
    ClassNN := ControlGetClassNN(ControlGetFocus(hwnd))
    }

If ClassNN != "DirectUIHWND2" {
    info := "ClassNN error in File Explorer! ExplorerTab_GetIdentity() aborted!"            "`n"
         .  "Title: "       WinGetTitle(hwnd)                                               "`n"
         .  "ahk_class: "   win_class                                                       "`n"
         .  "ahk_exe: "     WinGetProcessName(hwnd)                                         "`n"
         .  "ClassNN: "     ClassNN
    If IsSet(mytext)
        info .= "`n" "ClassNN was reset to 'DirectUIHWND2' once!"                           "`n"
             .  mytext
    ShowErrExit(A_ThisHotkey ":: " info)
    }

Try activeTab := ControlGetHwnd("ShellTabWindowClass1", hwnd)
Catch as err {
    SetTimer () => CatchError_show( CatchError_details(err, "Failed to get ExplorerTab_GetIdentity(hwnd)") ), -1   ; 1ms ; new thread
    Exit
    }

For w in Explorers {
    If (w.hwnd = hwnd) {
        IID_IShellBrowser   := "{000214E2-0000-0000-C000-000000000046}"
        shellBrowser        := ComObjQuery(w, IID_IShellBrowser, IID_IShellBrowser)
        ComCall(3, shellBrowser, "uint*", &thisTab := 0)
        If thisTab = activeTab
            Return w
        }
    }
Else  ; If `For` loop fails
    ShowErrExit(A_ThisHotkey ":: w.hwnd didn't match! ExplorerTab_GetIdentity() aborted!"                   "`n"
            .   "Title: "       WinGetTitle(hwnd)                                                           "`n"
            .   "ahk_class "    win_class "ahk_exe "     WinGetProcessName(hwnd)                            "`n"
            .   "For loop failed because of zero iterations."                                                   )
}

;--------
;    + ExplorerTab_Navigate
; modified from https://www.autohotkey.com/boards/viewtopic.php?p=4432#p4432

ExplorerTab_Navigate(newPath, hwnd := WinExist("File Explorer ahk_class CabinetWClass")) {
; File Explorer is specified here because navigating from other windows like "Control Panel" will use classic explorer and result in ugly artefacts in dark mode (e.g., vertical scrollbars)

identity    := ExplorerTab_GetIdentity(hwnd)
oldPath     := identity.Document.Folder.Self.Path

If oldPath != newPath     ; case insensitive match
    identity.Navigate("file:///" newPath)
Else ToolTipFn("ExplorerTab_Navigate() oldPath = newPath!", 2000) ; 2s
}

;--------
;    + ExplorerTab_GetFolderPath

ExplorerTab_GetFolderPath(hwnd := WinExist("A")) {
Return ExplorerTab_GetIdentity(hwnd).Document.Folder.Self.Path
}

;--------
;    + ExplorerTab_GetSelectionPath
; modified from https://www.autohotkey.com/boards/viewtopic.php?p=529074#p529074
; and https://www.autohotkey.com/boards/viewtopic.php?p=511944#p511944
; and https://www.autohotkey.com/boards/viewtopic.php?f=76&t=60403&p=255273#p255256

ExplorerTab_GetSelectionPath(hwnd := WinExist("A")) {

identity    := ExplorerTab_GetIdentity(hwnd).Document

pathContent := ""

For i, in identity.SelectedItems
    pathContent .= (A_Index > 1 ? "`n" : "") i.Path ; append each selected item and add a new line.
Else                                                ; For fails due to zero iterations
    pathContent .= identity.Folder.Self.Path

Return pathContent
}

;--------
;    + Explorer_ValidatePath
; pathType = (1) multiple folder/files, (2) single folder, (3) single file, (4) CLSID or (0) invalid mono/multi-line

Explorer_ValidatePath(pathContent) {
If InStr(pathContent, "`n") {                                               ; pathContent contains multiple folder/files
    Loop Parse pathContent, "`n", "`r" {
        If DirExist(A_LoopField) || FileExist(A_LoopField)
            invalid := 0
        Else {
            invalid := 1
            Break
            }
        }
    If invalid
        Return 0
    Else Return 1
    }
Else If DirExist(pathContent)                                               ; pathContent contains a single directory/folder path that is valid
    Return 2
Else If FileExist(pathContent)                                              ; pathContent contains a single file path that is valid
    Return 3
Else If RegExMatch(pathContent, "\{\w{8}-\w{4}-\w{4}-\w{4}-\w{12}\}") {     ; pathContent contains a CLSID
    ; e.g., Named folder path like {21EC2020-3AEA-1069-A2DD-08002B30309D} for `All Control Panel Items`
    Return 4
    }
Else Return 0                                                               ; pathContent is NOT valid
}

;--------
;    + Explorer_FocusAddressBar (unused)
; check focus of cursor in explorer before further action

Explorer_FocusAddressBar() {
Send "{F4}"
While ControlGetClassNN(ControlGetFocus("A")) != "Microsoft.UI.Content.DesktopChildSiteBridge1" {
    Sleep 100
    If A_Index > 5 {                                        ; = Sleep 500ms ; wait until focus is on address bar
        ToolTipFn("Failed to focus address bar!", 2000)     ; 2s
        Return 0
        }
    }
Send "A{Backspace}"
WinWait("PopupHost ahk_class Microsoft.UI.Content.PopupWindowSiteBridge",, 2) ; 2s - wait for drop down
Return 1
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
Else ShowErrExit(key ":: WallpaperPath_v4() is not valid!`nConverted String: " ConvString   "`n`n"
            .    "TranscodedImageCache:"                                                    "`n"
            .    regBinary                                                                          )
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
    SetTimer () => CatchError_show( CatchError_details(err, "nxtBackground() failed!") ), -1   ; 1ms ; new thread
}

;------------------------------------------------------------------------------
;  = File create, overwrite, append

;    + FileCreate_Overwrite
; modified from AutoHotkey.chm::/docs/lib/FileOpen.htm#ExWriteRead

FileCreate_Overwrite(FilePath := A_MyDocuments "\Test.txt" , FileContent := "Test", Encoding := "UTF-8") {

FileContent := Trim(FileContent, "`r`n`t`s`v`f")

If FileExist(FilePath) {                                    ; If file exists
    Try FileObj := FileOpen(FilePath, 1 + 4 , Encoding)     ; overwrite file contents
    ; 1 = Write: Creates a new file, overwriting any existing file
    ; 4 = Replace `r`n with `n when reading and `n with `r`n when writing

    Catch as err
        SetTimer () => CatchError_show( CatchError_details(err, "FileCreate_Overwrite() Can't open file for writing") ), -1                                 ; 1ms ; new thread
    Else {
        FileObj.Write(FileContent)
        FileObj.Close()
        }
    }
Else FileAppend FileContent, FilePath, "`n " Encoding       ; create new file
}

;------------------------------------------------------------------------------
;  = MsgBox Fn

;    + ShowErrExit

ShowErrExit(info) {

CatchError_show(info)

Global errorDetails := ""
Exit
}

;--------
;    + MsgBox_Custom

MsgBox_Custom(MsgText := "Text", MsgTitle := A_ScriptName, msgOption := 0, nButtons := 1, bName1 := "Button_Name", bName2?, bName3?) {  ; nButtons = 1/2/3 button options

MsgText .= "`n`n" "Buttons: // "

If nButtons = 1 { ; && isSet(bName1)
    msgOption   += 0                                        ; OK button
    MsgText     .= "OK ("       bName1 ")"  " //"
}

Else If nButtons = 2 && isSet(bName2) {
    msgOption   += 4                                        ; YesNo button
    MsgText     .= "Yes ("      bName1 ")"  " // "
                .  "No ("       bName2 ")"  " //"
}

Else If nButtons = 3 && isSet(bName2) && isSet(bName3) {
    msgOption   += 3                                        ; YesNoCancel button
    MsgText     .= "Yes ("      bName1 ")"  " // "
                .  "No ("       bName2 ")"  " // "
                .  "Cancel ("   bName3 ")"  " //"
}

Else ShowErrExit(A_ThisHotkey ":: MsgBox_Custom()"                    "`n"
            .    "Error! Check options and button number and names"       )

RandomNo := Random(1, 194161877)
SetTimer () => MsgBox_ChangeButtonText(RandomNo, MsgTitle, nButtons, bName1, bName2?, bName3?), -100 ; 100ms ; new thread
Return MsgBox(MsgText, RandomNo, msgOption + 262144)        ; 262144 Always-on-top
}

;--------
;    + MsgBox_ChangeButtonText

MsgBox_ChangeButtonText(RandomNo, MsgTitle, nButtons, bName1, bName2?, bName3?) {

errorDetails := ""
If RandomNo = 0 {
    If WinWait(MsgTitle " ahk_class #32770",, 2)            ; 2s
        MsgBox_setButtonNames()
    Else errorDetails .= ":: MsgBox_ChangeButtonText failed! MsgTitle - " MsgTitle ": Timed out!"
    }
Else If WinWait(RandomNo " ahk_class #32770",, 2) {         ; 2s
    WinSetTitle MsgTitle
    MsgBox_setButtonNames()
    }
Else errorDetails .= ":: MsgBox_ChangeButtonText failed! RandomNo - " RandomNo ": Timed out!"

If errorDetails != ""
    ToolTipFn(A_ThisHotkey "`n" errorDetails, 2000)         ; 2s
Else Return

;--------
;    + MsgBox_setButtonNames (enclosed Fn)

    MsgBox_setButtonNames() {
    ControlSetText bName1, "Button1"

    If isSet(bName2)
        ControlSetText bName2, "Button2"

    If isSet(bName3)
        ControlSetText bName3, "Button3"
    }
}

;------------------------------------------------------------------------------
;  = Compare Arrays & Remove Duplicates

;    + arrCompare_RemoveDuplicates

arrCompare_RemoveDuplicates(arr1, arr2, ActOn := 2) {   ; Returns [arr1, arr2, count := 0, errorMsg := 0]

Position := ""

If ActOn = 1 {
    Loop arr2.Length {                                  ; If 1 or more values in `done` array
        present := arr2[A_Index]
        Loop arr1.Length {                              ; remove those values from `found` array if present
            If arr1[A_Index] == present                 ; case-sensitive
                Position .= A_Index ","
            }
        Else arrCompare_error("arrCompare_RemoveDuplicates() First array length is Zero!")
        }
    Else arrCompare_error("arrCompare_RemoveDuplicates() Second array length is Zero!")
    }

Else { ; ActOn = 2
    Loop arr1.Length {                                  ; If 1 or more values in `done` array
        present := arr1[A_Index]
        Loop arr2.Length {                              ; remove those values from `found` array if present
            If arr2[A_Index] == present                 ; case-sensitive
                Position .= A_Index "`n"
            }
        Else arrCompare_error("arrCompare_RemoveDuplicates() Second array length is Zero!")
        }
    Else arrCompare_error("arrCompare_RemoveDuplicates() First array length is Zero!")
    }

Return arrCompare_results()

;--------
;    + arrCompare_error (enclosed Fn)

    arrCompare_error(errorMsg) {
    ; ToolTipFn(A_ThisHotkey "::" errorMsg, 2000)         ; 2s
    Return [arr1, arr2, 0, errorMsg]                    ; action count is zero
    }

;--------
;    + arrCompare_results (enclosed Fn)

    arrCompare_results() {
    If Position = "" {
        ; ToolTipFn(A_ThisHotkey "::arrCompare_RemoveDuplicates() success! No duplicates found!", 2000)       ; 2s
        Return [arr1, arr2, 0, 0]                       ; count zero, no errorMsg
        }
    Else {
        count := 0
        Position := Sort(Position, "D, U R ")
        Loop Parse Position "," "`n`r " {
            If IsNumber(A_Loopfield) {
                count++
                If ActOn = 1
                    arr1.RemoveAt(A_Loopfield)
                Else arr2.RemoveAt(A_Loopfield)
                }
            }
        ; ToolTipFn(A_ThisHotkey "::arrCompare_RemoveDuplicates() success! Duplicate Count: " count, 2000)    ; 2s
        Return [arr1, arr2, count, 0]                   ; no errorMsg
        }
    }

}

;------------------------------------------------------------------------------
;  = #HotIf functions

;    + checkCalcView

checkCalcView() {

DetectHiddenText False  ; 0
WinText := WinGetText("ahk_class CalcFrame")
DetectHiddenText True   ; 1

If InStr(WinText, "Select the type of unit you want to convert")
    CalcView := 1       ; Unit conversion view
Else If InStr(WinText, "Select the date calculation you want")
    CalcView := 2       ; Date Calculation view
Else If InStr(WinText, "Down payment")  || InStr(WinText, "Purchase price")
    CalcView := 3       ; Worksheets > Mortgage
Else If InStr(WinText, "Lease value")   || InStr(WinText, "Residual value")
    CalcView := 4       ; Worksheets > Vehicle lease
Else If InStr(WinText, "Fuel ")         && not RegExMatch(WinText, "liters|kilometers")
    CalcView := 5       ; Worksheets > Fuel economy (mpg)
Else If InStr(WinText, "Fuel ")         && not RegExMatch(WinText, "gallons|miles")
    CalcView := 6       ; Worksheets > Fuel economy (L/100 km)
Else CalcView := 0      ; Basic view

Return CalcView
}

;--------
;    + MarkletFn

MarkletFn(key) {
; key must match `Keyword` field of saved bookmarklet in Firefox library
; check here for more details - https://github.com/xypha/Bookmarklets#introduction-to-bookmarklets

Send "^l"                  ; focus address bar
Sleep 250                  ; wait for focus
Send key "{Enter}"         ; Run bookmarklet
}

;--------
;    + CloseIncrSearch

CloseIncrSearch(key) {
DetectHiddenText False                                  ; 0 ; only check visible text
If WinActive("ahk_class Notepad++", "Find:") && ControlGetClassNN(ControlGetFocus()) = "Edit1" {
; If incremental search is present (i.e. "Find:" text is visible) and has keyboard focus
    Send "^a"                                           ; Ctrl + A ; select all text typed in search field     ~~ comment out if not needed
    Send "^c"                                           ; Ctrl + C ; copy search term to clipboard             ~~ comment out if not needed
    Send "{Esc}"                                        ; close Incremental Search
    Sleep 100                                           ; wait for closure/focus main window                   ~~ comment out if not needed
    }
DetectHiddenText True                                   ; 1 ; restore default
Send RegExReplace(key, "([a-zA-Z]{2,})" , "{$1}")       ; add {} around words like Tab, Up…
; Sleep 100                                               ; wait for dialogue to open and focus           ~~ uncomment if ^v needed
; Send "^v"                                               ; Ctrl + V ; paste search term from clipboard   ~~ uncomment if needed
}

;------------------------------------------------------------------------------
;  = Windows Registry

;    + RegJump
; modified from https://gist.github.com/raveren/bac5196d2063665d2154#file-aio-ahk-L811

RegJump(regPath := A_Clipboard) {

; remove leading Computer
regPath := StrReplace(regPath, "Computer\")

; if input contains multiple lines, create array and Loop to find a valid regPath
regArray := StrSplit(regPath, "`n", "`r")

regFound := 0
Loop regArray.Length {
    If regArray[A_Index] ~= "^HKEY_|^HKCU\\|^HKLM\\|^HKU\\|^HKCC\\" { ; escape \ with \
        regPath     := regArray[A_Index]
        regFound    := 1
        Break
        }
    }

If regFound = 0
    ShowErrExit(A_ThisHotkey ":: RegJump() failed! regArray match not found!"   "`n"
            .   "Path: " regPath                                                    )
; Else regFound = 1 ; valid regPath is found, so proceed

; remove trailing "\" if present ; detect Space too using RegEx?
If SubStr(regPath, -1) = "\"            ; -1 extracts the last character
    regPath := SubStr(regPath, 1, -1)   ; a negative Length to omit that many characters from the end

; check if regPath is valid
If not Registry_ValidPath(regPath) ; Returns 0 = failure/error
    ShowErrExit(A_ThisHotkey ":: RegJump() failed! regPath is not valid!"       "`n"
            .   "Path: " regPath                                                    )
; Else 1 ; Returns 1 = success ; regPath is a valid registry path

; Must close Regedit so that next time it opens the target key is selected
If WinExist("Registry Editor")
    WinClose "Registry Editor" ; alternative WinKill

If not WinWaitClose("Registry Editor",, 2) ; 2s
    ShowErrExit(A_ThisHotkey ":: RegJump() aborted because WinClose timed out!" "`n"
            .   "Path: " regPath                                                    )
; Else ; window is closed, proceed as below

; Extract RootKey part of supplied registry path HKEY_CURRENT_USER\Software\Microsoft\Windows\Current Version
Loop Parse, regPath, "\" {
    RootKey := A_LoopField
    Break
}

; Now convert RootKey to standard long format
If not InStr(RootKey, "HKEY_") { ; if short form, convert to long form
    If RootKey = "HKCR"
        regPath := StrReplace(regPath, RootKey, "HKEY_CLASSES_ROOT"     ,,, 1)
    Else If RootKey = "HKCU"
        regPath := StrReplace(regPath, RootKey, "HKEY_CURRENT_USER"     ,,, 1)
    Else If RootKey = "HKLM"
        regPath := StrReplace(regPath, RootKey, "HKEY_LOCAL_MACHINE"    ,,, 1)
    Else If RootKey = "HKU"
        regPath := StrReplace(regPath, RootKey, "HKEY_USERS"            ,,, 1)
    Else ; RootKey = "HKCC"
        regPath := StrReplace(regPath, RootKey, "HKEY_CURRENT_CONFIG"   ,,, 1)
}

; Make target key the last selected key, which is the selected key next time Regedit runs
RegWrite regPath, "REG_SZ", "HKCU\Software\Microsoft\Windows\CurrentVersion\Applets\Regedit", "LastKey"

Run "Regedit.exe"
}

/* physical locations - https://superuser.com/questions/111311/where-are-registry-files-stored-in-windows
HKEY_LOCAL_MACHINE\SAM                                  : %SystemRoot%\System32\config\sam
HKEY_LOCAL_MACHINE\SECURITY                             : %SystemRoot%\System32\config\security
HKEY_LOCAL_MACHINE\SOFTWARE                             : %SystemRoot%\System32\config\software
HKEY_LOCAL_MACHINE\SYSTEM                               : %SystemRoot%\System32\config\system
HKEY_USERS.DEFAULT                                      : %SystemRoot%\System32\config\default
file associations (per user - HKCR\Software\Classes)    : %UserProfile%\Local Settings\Application Data\Microsoft\Windows\UsrClass.dat
per user settings                                       : %UserProfile%\NTUSER.DAT
*/

;--------
;    + Registry_ValidPath

Registry_ValidPath(regPath) {
Try {
    Loop Reg, regPath, "R KV" ; Recursively retrieve keys and values
        Return 1    ; True
    Else            ; The loop had zero iterations
        Return 0    ; False
    }
Catch
    Return 0        ; False
}

;------------------------------------------------------------------------------
;  = Control Panel Tools

;    + ControlPanel_menuFn

ControlPanel_menuFn() {
If IsSet(ControlPanel_menu) {
    ControlPanel_menu.Show
    Return
    }
; Else ; Continue as below

Global ControlPanel_menu := Menu()      ; starts building a pop-up menu
; ControlPanel_menu.Delete()            ; replaced by If IsSet() ; deletes previously built pop-up menu, if any, and then starts adding items

; add menu items
ControlPanel_menu.Add("&1  → Control Panel"                 , ControlPanel_selection )
ControlPanel_menu.Add("&2  → Installed Apps"                , ControlPanel_selection )
ControlPanel_menu.Add("&3  → Add/Remove Programs (Legacy)"  , ControlPanel_selection )
ControlPanel_menu.Add("&4  → Defragment Interface"          , ControlPanel_selection )
ControlPanel_menu.Add("&5  → Services"                      , ControlPanel_selection )
ControlPanel_menu.Add("&6  → Sound Mixer (Legacy)"          , ControlPanel_selection )
ControlPanel_menu.Add("&7  → Registry Editor"               , ControlPanel_selection )
ControlPanel_menu.Add("&8  → Resource Monitor"              , ControlPanel_selection )
ControlPanel_menu.Add("&9  → Windows Update"                , ControlPanel_selection )
ControlPanel_menu.Add("&0  → Task Scheduler"                , ControlPanel_selection )

; add menu icons
MenuIcons_add(ControlPanel_menu, 10)

; show pop-up menu
ControlPanel_menu.Show
}

;--------
;    + ControlPanel_selection

ControlPanel_selection(ItemName, ItemPos, MyMenu) {
If ItemPos = 1
    ComObject("shell.application").ControlPanelItem("control")      ; Control Panel
If ItemPos = 2
    Run 'explorer.exe "ms-settings:appsfeatures"'                   ; Installed Apps ; Modern Add/Remove Programs
If ItemPos = 3
    ComObject("shell.application").ControlPanelItem("appwiz.cpl")   ; Add/Remove Programs ; Legacy Control Panel
If ItemPos = 4
    ComObject("shell.application").ControlPanelItem("dfrgui")       ; Defragment Interface
If ItemPos = 5
    ComObject("shell.application").ControlPanelItem("services.msc") ; Services
If ItemPos = 6
    ComObject("shell.application").ControlPanelItem("sndvol")       ; Sound Mixer (Legacy)
If ItemPos = 7
    ComObject("shell.application").ControlPanelItem("regedit")      ; Registry Editor
If ItemPos = 8
    ComObject("shell.application").ControlPanelItem("resmon.exe")   ; Resource Monitor
If ItemPos = 9
    Run 'explorer.exe "ms-settings:windowsupdate"'                  ; Windows Update
If ItemPos = 10
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
shell:::{ED834ED6-4B5A-4bfe-8F11-A626DCB6A921} -Microsoft.Personalization\pageWallpaper     ; old windows wallpaper customizer from windows7
ComObject("shell.application").ControlPanelItem("winver")                                   ; Windows version
Run 'SnippingTool.exe'                                                                      ; Snipping Tool ; alternate ; Opens modern app
Run 'rundll32 sysdm.cpl`,EditEnvironmentVariables'                                          ; Environmental Variables
ComObject("shell.application").ControlPanelItem("powercfg.cpl")                             ; Power Configuration ; opens Power Options
ComObject("shell.application").ControlPanelItem("msinfo32")                                 ; System Information
ComObject("shell.application").ControlPanelItem("timedate.cpl")                             ; Date and Time Properties
ComObject("shell.application").ControlPanelItem("ncpa.cpl")                                 ; Network Connections
ComObject("shell.application").ControlPanelItem("mmsys.cpl")                                ; Sounds and Audio ; Opens old Sound panel - Playback, Recording, Sounds, Communications
ComObject("shell.application").ControlPanelItem("dcomcnfg")                                 ; Component Services
ComObject("shell.application").ControlPanelItem("gpedit.msc")                               ; Group Policy Editor ; N/A in Home
ComObject("shell.application").ControlPanelItem("iexplore")                                 ; Internet Explorer ; Opens Edge browser
ComObject("shell.application").ControlPanelItem("inetcpl.cpl")                              ; Internet Properties
ComObject("shell.application").ControlPanelItem("secpol.msc")                               ; Local Security Settings ; N/A in Home
ComObject("shell.application").ControlPanelItem("lusrmgr.msc")                              ; Local Users and Groups ; N/A in Win10 & later?
ComObject("shell.application").ControlPanelItem("logoff")                                   ; Logs You Out Of Windows
ComObject("shell.application").ControlPanelItem("main.cpl")                                 ; Mouse Properties
ComObject("shell.application").ControlPanelItem("perfmon.msc")                              ; Performance Monitor
ComObject("shell.application").ControlPanelItem("intl.cpl")                                 ; Regional Settings
ComObject("shell.application").ControlPanelItem("mstsc")                                    ; Remote Desktop ; N/A in Home
ComObject("shell.application").ControlPanelItem("wscui.cpl")                                ; Security and Maintenance
ComObject("shell.application").ControlPanelItem("fsmgmt.msc")                               ; Shared Folders/MMC
ComObject("shell.application").ControlPanelItem("shutdown")                                 ; Shuts Down Windows
ComObject("shell.application").ControlPanelItem("StikyNot")                                 ; Sticky Note ; N/A
ComObject("shell.application").ControlPanelItem("msconfig")                                 ; System Configuration Utility
ComObject("shell.application").ControlPanelItem("sysdm.cpl")                                ; System Properties
ComObject("shell.application").ControlPanelItem("taskmgr")                                  ; Task Manager
ComObject("shell.application").ControlPanelItem("netplwiz")                                 ; User Accounts
ComObject("shell.application").ControlPanelItem("utilman")                                  ; Modern Settings App > Accessibility
ComObject("shell.application").ControlPanelItem("firewall.cpl")                             ; Windows Defender Firewall
ComObject("shell.application").ControlPanelItem("wf.msc")                                   ; Windows Defender Firewall with Advanced Security
ComObject("shell.application").ControlPanelItem("wmimgmt.msc")                              ; Windows Management Instrumentation (WMI)
ComObject("shell.application").ControlPanelItem("wuapp")                                    ; Windows Update App Manager ; N/A
ComObject("shell.application").ControlPanelItem("write")                                    ; WordPad ; N/A
ComObject("shell.application").ShutdownWindows()                                            ; Shutdown Menu

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
Shell:AppsFolder    ; shortcuts to all apps in start menu
shell:start menu    ; location of start menu shortcuts
PostMessage 0x0111, 65305,,, "C:\YourScript.ahk ahk_class AutoHotkey" ; Suspend, Toggle
PostMessage 0x0111, 65306,,, "ScriptFileName.ahk - AutoHotkey" ; Pause, Toggle
PostMessage 0x0111, 65303,,, "ScriptFileName.ahk - AutoHotkey"  ; Reload.

restart your video drivers by pressing the key combination Win + Ctrl + Shift + B -- https://www.makeuseof.com/tag/hidden-key-combo-frozen-computer/
Classic Control Panel = control.exe
more? Check jeeswg's Explorer tutorial - https://www.autohotkey.com/boards/viewtopic.php?p=148121#p148121

AppsFolder Shell folder name registry key (source: https://www.techrepublic.com/article/how-to-use-the-shell-command-to-view-all-your-applications-in-file-explorer/)
HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderDescriptions
KNOWNFOLDERID - https://learn.microsoft.com/en-us/windows/win32/shell/

A_Downloads := RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "{374DE290-123F-4565-9164-39C4925E467B}")
A_Downloads := ComObject('Shell.Application').NameSpace('Shell:Downloads').Self.Path ; source: https://www.autohotkey.com/boards/viewtopic.php?p=285131#p285131
*/

;------------------------------------------------------------------------------
; * Test

:*:test++:: {
MsgBox ThisHotkey ":: Not assigned!",, 262144 ; 262144 Always-on-top

}

; End of script code
