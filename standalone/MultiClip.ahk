; https://github.com/xypha/AHK-v2-scripts/edit/main/standalone/MultiClip.ahk
; Last updated 2024.11.29
; This is (mostly) feature complete. There won't be any major changes henceforth.
; For a more feature rich and frankly awesome ahk clipboard manager, visit https://github.com/mikeyww/mwClipboard/ 

; Visit AutoHotkey (AHK) version 2 (v2) help for information - https://www.autohotkey.com/docs/v2/
; Search for commands/functions used in this script by using Ctrl + F on the AutoHotkey help webpage - https://www.autohotkey.com/docs/v2/lib/

; comments begin with semi-colon ";" at start of line or space or after code in middle of line
; comments can also be enclosed by `/* */`, like this - /* comment text */
; and these two methods can be combined too :)

;   /* AHK v2 MultiClip - CONTENTS */
; Settings
; Auto-execute
;  = AHK Dark Mode
;  = Initialise ClipArr
;  = Initialise ClipArr hotstrings
;  = Tray Icon
;  = End auto-execute
; Hotkeys
;  = Check & Reload AHK
; Hotstrings
;  = ClipArr keys
;  = ClipArr testing
; User-defined functions
;  = MyNotification
;    + MyNotificationGui
;    + EndMyNotif
;  = AHK Dark Mode Fn
;    + ahkDarkMenu
;  = MultiClip ClipArr
;    + ClipChanged
;    + InsertInClipArr
;    + ClipArr ToolTipFn
;    + SaveClipArr
;    + PasteVStrings
;    + PasteCStrings
;  = MultiClip ClipMenu
;    + ClipMenuFn
;    + ClipTrim
;    + SendClipFn
;  = Paste instead of Send
;    + PasteThis
;    + Paste_via_clipboard
;    + RestoreClip
;  = ToolTip function
;    + ToolTipFn

;------------------------------------------------------------------------------
; Settings
; Start of script code

#Requires AutoHotkey v2.0
#SingleInstance force
KeyHistory 500 ; Max 500

;------------------------------------------------------------------------------
; Auto-execute
; This section should always be at the top of your script

AHKname := "AHK v2 No-2 MultiClip v4.12"

; Show notification with parameters - text; duration in milliseconds; position on screen: xAxis, yAxis; timeout by - timer (1) or sleep (0)
MyNotificationGui("Loading " AHKname, 10000, 1550, 945, 1) ; 10000ms = 10 seconds, position bottom right corner (x-axis 1550 y-axis 985) on 1920×1080 display resolution; use timer

;--------
;  = AHK Dark Mode

; check windows registry to see if dark mode is enabled
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
; manually comment out above lines if dark mode is NOT enabled because "#Include cannot be executed conditionally"
; disable dark mode commands "MyNotification.AddText" and "MyNotification.BackColor" in `MyNotificationGui` function

;--------
;  = Initialise ClipArr

Global LimitClipArr := 25
; Limit the number of slots to 25 to match Win + V clipboard history slots ; customise limit to your needs
; Note:     Higher the number, higher the resource usage and slower the performance/response
; WARNING:  Removing the variable entirely may cause infinite loop / app hang

Global ClipArrFile := A_MyDocuments "\ClipArrFile.txt"
; ClipArrFile.txt is saved in default path of AHK's built-in variable: A_MyDocuments
; A_MyDocuments is the full path and name of the current user's "My Documents" folder. Usually corresponds to "C:\Users\<UserName>\Documents" (the final backslash is not included in the variable)
; txt file used instead of ini because 'values longer than 65,535 characters are likely to yield inconsistent results' when using IniRead, IniWrite commands

Global delim := "~•~"
; ~ U+007E TILDE
; • U+2022 BULLET : black small circle
; use a unique string because if an array-slot contains this delimiter by accident, saving and loading array from file will cause errors
; Recommendation: 3 or more characters, preferably symbols with one or more Unicode characters that are difficult to type on standard keyboard. For suggestions, look here - https://stackoverflow.com/questions/492090/least-used-delimiter-character-in-normal-text-ascii-128

Global ClipArr := [] ; set Global variable and assign empty array

; Load array from file if file exists - inspired by https://www.autohotkey.com/boards/viewtopic.php?p=341809#p341809
; `Try` command is used to prevent AutoHotkey from throwing error msg in case file is absent or not in correct path
Try ClipArr := StrSplit(FileRead(ClipArrFile, "UTF-8"), delim, , LimitClipArr)
Catch ; or Else load default values on start - 25 slots containing alphanumerical text
    ClipArr := ["Slot 1 Shortcut 1",
                "Slot 2 Shortcut 2",
                "Slot 3 Shortcut 3",
                "Slot 4 Shortcut 4",
                "Slot 5 Shortcut 5",
                "Slot 6 Shortcut 6",
                "Slot 7 Shortcut 7",
                "Slot 8 Shortcut 8",
                "Slot 9 Shortcut 9",
                "Slot 10 Shortcut 0",
                "Slot 11 Shortcut Q",
                "Slot 12 Shortcut w",
                "Slot 13 Shortcut e",
                "Slot 14 Shortcut r",
                "Slot 15 Shortcut t",
                "Slot 16 Shortcut Y",
                "Slot 17 Shortcut u",
                "Slot 18 Shortcut i",
                "Slot 19 Shortcut o",
                "Slot 20 Shortcut P",
                "Slot 21 Shortcut a",
                "Slot 22 Shortcut s",
                "Slot 23 Shortcut d",
                "Slot 24 Shortcut f",
                "Slot 25 Shortcut G"]

; run function whenever clipboard is changed such as when Ctrl + x (Cut) or Ctrl + c (Copy) is pressed
; or when clipboard is altered by other apps/programs
OnClipboardChange ClipChanged

; add current clipboard contents to first clipboard slot in ClipArr on start
InsertInClipArr(A_Clipboard, 1) ; onstart = 1

; save `ClipArr` contents to `ClipArrFile.txt` when the script exits
; except when it is killed by something like "End Task" via Taskbar, Task Manager or similar
OnExit SaveClipArr

;--------
;  = Initialise ClipArr hotstrings

PasteVStrings(LimitClipArr)   ; User-defined function creates serial hotstrings
PasteCStrings(LimitClipArr)

;--------
;  = Tray Icon

path_to_TrayIcon := A_ScriptDir "\icons\Tray\2-512.ico"
; Icon source: https://www.iconsdb.com/caribbean-blue-icons/2-icon.html     ; CC License, see credits.md
; I like to number scripts 1, 2, 3... and link the scripts to Numpad shortcuts for easy editing -- see section on "Check & Reload AHK" below

Try TraySetIcon path_to_TrayIcon
Catch as err
    SetTimer () => MsgBox("TraySetIcon failed!`n" err.Message "`nPath: " path_to_TrayIcon,, 262144), -100 ; 100ms ; new thread ; 262144 = Always-on-top

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

!Numpad2:: { ; Alt + Numpad2 keys pressed together
ListLines
If WinWait(A_ScriptFullPath " - AutoHotkey v" A_AhkVersion,, 3) ; 3s timeout ; wait for ListLines window to open
    WinMaximize
}

^!Numpad2:: {                                       ; Ctrl + Alt + Numpad2 keys pressed together
MyNotificationGui("Updating " AHKname,,, 945, 0)    ; 500ms ; use Sleep because reload cancels timers
Reload
}

;------------------------------------------------------------------------------
; Hotstrings

;  = ClipArr keys

:?*x:v0+::PasteV(10) ; same as v10+ ; pastes value in slot #10 in ClipArr ; default value "j10"

:?*x:c1+::Send "{Raw}" ClipArr[1] ; Send first entry in raw mode, useful when Ctrl + V is disabled such as on banking sites

:?*x:c0+::PasteC(10) ; same as c10+

:?*x:c++::ClipMenuFn(SendClipFn) ; show ClipMenu

;--------
;  = ClipArr testing

; test MultiClip function
:*:testclip+:: {

; save current array contents to file
SaveClipArr()
    ; If script is reloaded after test, saved file will be deleted. In this case, restore array contents by
    ; (a) Exiting this script (b) restoring deleted file from recycle bin (c) running this script again

A_Clipboard := "Slot 1 Shortcut 1"
Global ClipArr := ["Slot 1 Shortcut 1",
                "Slot 2 Shortcut 2",
                "Slot 3 Shortcut 3",
                "Slot 4 Shortcut 4",
                "Slot 5 Shortcut 5",
                "Slot 6 Shortcut 6",
                "Slot 7 Shortcut 7",
                "Slot 8 Shortcut 8",
                "Slot 9 Shortcut 9",
                "Slot 10 Shortcut 0",
                "Slot 11 Shortcut Q",
                "Slot 12 Shortcut w",
                "Slot 13 Shortcut e",
                "Slot 14 Shortcut r",
                "Slot 15 Shortcut t",
                "Slot 16 Shortcut Y",
                "Slot 17 Shortcut u",
                "Slot 18 Shortcut i",
                "Slot 19 Shortcut o",
                "Slot 20 Shortcut P",
                "Slot 21 Shortcut a",
                "Slot 22 Shortcut s",
                "Slot 23 Shortcut d",
                "Slot 24 Shortcut f",
                "Slot 25 Shortcut G"]
ClipMenuFn(SendClipFn)  ; show menu - ClipMenu
}

; restore old ClipArr contents from file after testing or from previous file when needed
:*x:restoreclip+:: {
Global ClipArr := ""
ClipArr := StrSplit(FileRead(ClipArrFile, "UTF-8"), delim, , LimitClipArr)  ; restore from file
ClipMenuFn(SendClipFn)                                                      ; show menu - ClipMenu
}

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
;  = AHK Dark Mode Fn

;    + ahkDarkMenu
/* primary source: https://stackoverflow.com/a/58547831/894589
   with modifications by
     * lexikos https://www.autohotkey.com/boards/viewtopic.php?p=426482#p426482
     * mcd https://www.autohotkey.com/boards/viewtopic.php?p=511756#p511756
*/

ahkDarkMenu() {
    Static uxtheme := DllCall("GetModuleHandle", "str", "uxtheme", "ptr")
    Static SetPreferredAppMode := DllCall("GetProcAddress", "ptr", uxtheme, "ptr", 135, "ptr")
    Static FlushMenuThemes := DllCall("GetProcAddress", "ptr", uxtheme, "ptr", 136, "ptr")

    DllCall(SetPreferredAppMode, "int", 1) ; 0 = Default, 1 = AllowDark, 2 = ForceDark, 3 = ForceLight, 4=Max
    DllCall(FlushMenuThemes)
}

;------------------------------------------------------------------------------
;  = MultiClip ClipArr
; Sources - https://www.autohotkey.com/boards/viewtopic.php?p=326827#p326827
; and MultiClip v1 - https://www.autohotkey.com/boards/viewtopic.php?p=332658#p332658
; MultiClip v1 used or modified AHK v1 code from https://autohotkey.com/board/topic/4567-clipstep-step-through-multiple-clipboards-using-ctrl-x-c-v/
; and https://geekdrop.com/content/super-handy-autohotkey-ahk-script-to-change-the-case-of-text-in-line-or-wrap-text-in-quotes

;    + ClipChanged

ClipChanged(DataType) {

If DataType = 0 {                                               ; 0 Clipboard is now empty
    ; ToolTipFn("DataType: 0 - Clipboard is now empty", 1000)   ; 1s
    Exit
    }

Else If DataType = 2 {                                  ; 2 Clipboard contains something entirely non-text such as a picture
    ToolTipFn("DataType: 2 - Non-text copied", 1000)    ; 1s
    Exit
    }

; Else DataType = 1
; Clipboard contains text (including files copied from Windows File Explorer)
; check and add to clipArr (in case of files, file path is copied to clipArr)
Else InsertInClipArr(A_Clipboard)
}

;--------
;    + InsertInClipArr

InsertInClipArr(text, onStart := 0) {

Cliptemp := StrReplace(text,"`r`n","`n")        ; fix for SendInput sending Windows line-breaks

If RegExMatch(Cliptemp,"^\s+$") {               ; don't insert empty strings into clipArr
    If onStart != 1                             ; If NOT on startup/reload
        ToolTipFn("[~ Only \s ~]", 2000)        ; 2s ; show alert instead
    Exit
    }

Cliptemp := RegExReplace(Cliptemp,"^\s+|\s+$")  ; remove leading/trailing \s = [\r\n\t\f\v ]

; use Loop to check if Cliptemp is already in an array. If found, remove it and retrieve its `Index`
Loop LimitClipArr {
    If Cliptemp == ClipArr[A_Index] {
        ClipArr.RemoveAt(A_Index)
        FoundClip := A_Index
        Break
        }
    }
ClipArr.InsertAt(1, Cliptemp)   ; insert current clipboard contents in the first slot
ClipArr.Length := LimitClipArr  ; reset number of slots to previously defined limit
If IsSet(FoundClip) and onStart = 1 and FoundClip = 1
    Return
Else ClipArrToolTipFn()
}

;--------
;    + ClipArr ToolTipFn
; clipboard change alert tooltip

ClipArrToolTipFn() {
If StrLen(ClipArr[1]) > 1000                                ; trim If more than 1000 characters
    ToolTipFn(SubStr(ClipArr[1], 1, 1000) "`n… and more")   ; 500ms
Else ToolTipFn(ClipArr[1])                                  ; 500ms
}

;--------
;    + SaveClipArr
; use * because OnExit Callback accepts two parameters

SaveClipArr(*) {
Result := ""

; save current ClipArr contents to variable
Loop (ClipArr.Length > LimitClipArr ? LimitClipArr : ClipArr.Length)
    Result .= ClipArr[A_Index] delim

; remove trailing delim to prevent ClipArr.Length from exceeding LimitClipArr on restoration of ClipArr
Result := SubStr(Result, 1, StrLen(delim) * -1)

Try FileRecycle ClipArrFile                     ; send old file to recycle bin if one exists
    ; old clipboard contents can be retrieved by restoring ClipArrFile from recycle bin
    ; alternatively, use `FileDelete` command to delete permanently
FileAppend Result, ClipArrFile, "`n UTF-8"      ; create new file and save current clipArr contents
}

;--------
;    + PasteVStrings

PasteVStrings(number) {
Loop number
    Hotstring(":?*x:v" A_Index "+", PasteV)
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

PasteV(hk) {
RegExMatch(hk, "\d+", &SubPat)
Try PasteThis(ClipArr[SubPat[]])
; Try Send ClipArr[SubPat[]] ; alternative
}

;--------
;    + PasteCStrings

PasteCStrings(number) {
Loop number {
    If A_Index = 1  ; do not create c1+ hotstring, already assigned to "{Raw}" ClipArr[1]
        Continue    ; = same as saying "skip"
    Hotstring(":?*x:c" A_Index "+", PasteC)
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

PasteC(hk) {
RegExMatch(hk, "\d+", &SubPat)
PasteAll(SubPat[])
}

PasteAll(hkey) {
Loop hkey {
    Try clipVar := ClipArr[A_Index]
    Catch IndexError
        Result .= "`n"
    Else Result .= clipVar "`n"
    }
PasteThis(RegExReplace(Result,"^\n+|\n+$")) ; remove leading/trailing LF
}

;------------------------------------------------------------------------------
;  = MultiClip ClipMenu

;    + ClipMenuFn

ClipMenuFn(FnName) {

; create array to store shortcuts
ClipShortcuts := StrSplit("1234567890QwertYuioPasdfG")

/* ; shortcuts for 25 slots consist of numbers from number row, and letters from the rows below it in a QUERTY keyboard
; Customise the shortcut characters and their order by altering the characters in `ClipShortcuts` variable as needed
Slot 1 Shortcut 1 ; number row
Slot 2 Shortcut 2
Slot 3 Shortcut 3
Slot 4 Shortcut 4
Slot 5 Shortcut 5
Slot 6 Shortcut 6
Slot 7 Shortcut 7
Slot 8 Shortcut 8
Slot 9 Shortcut 9
Slot 10 Shortcut 0
Slot 11 Shortcut Q ; 1st letter row
Slot 12 Shortcut w
Slot 13 Shortcut e
Slot 14 Shortcut r
Slot 15 Shortcut t
Slot 16 Shortcut Y
Slot 17 Shortcut u
Slot 18 Shortcut i
Slot 19 Shortcut o
Slot 20 Shortcut P
Slot 21 Shortcut a ; 2nd letter row
Slot 22 Shortcut s
Slot 23 Shortcut d
Slot 24 Shortcut f
Slot 25 Shortcut G
; QYPG are capitals because selection underline causes confusion with other letters like o or v
; pressing shift + letter is not necessary because shortcuts are NOT case-sensitive
*/

Global ClipMenu := Menu()
ClipMenu.Delete

LoopNo := ClipArr.Length > LimitClipArr ? LimitClipArr : ClipArr.Length

; populate slots
Loop LoopNo {
    ClipMenu.Add("&" ClipShortcuts[A_Index] "  → " ClipTrim(A_Index), FnName)
    ; When the menu is displayed, a slot can be selected by pressing the key corresponding to the character preceded by the ampersand (&)
    ; These selection shortcuts correspond to the number/alphabet/symbol before `→` and are obtained from ClipShortcuts array
    ; When the menu is displayed, shortcuts are usually underlined, but sometimes don't appear when some symbols are used
    }

; Set icons for each menu item
Loop LoopNo {
    Try ClipMenu.SetIcon(A_Index "&", A_ScriptDir "\Icons\ClipMenu\" A_Index "-20.jpg", , 20)
    ; Icon Source: Calendar by Kalash - CC BY 4.0 - https://icon-icons.com/pack/Calendar/4173
    ; Icons were cropped using https://bulkimagecrop.com/ ; and converted to jpg and resized using mspaint (classic)
    ; `Try` command is used to prevent AutoHotkey from throwing error msgs in case icon files are absent or not in correct path.
    ; WARNING: Using icons in menu may slow performance.
    ; A slight delay between menu request and display may be noticeably present on some systems (especially in low-end ones like mine; probably to resize/rescale icons).
    ; This is normal and expected. Comment out the `ClipMenu.SetIcon` line if this is not acceptable.
    ; Default size is 16. Increased to 20 because icon numbers are not clear. Please let me know if you find icons that look better with `ahkDarkMenu` by creating an Issue
    Catch {
        icon_error := "Yes"
        Break
        }
    }

If isSet(icon_error)
    ClipMenu.Add("&// SetIcon command failed! // Path: " A_ScriptDir "\Icons\ClipMenu\", ClipMenu_icon_error)

; show pop-up menu
ClipMenu.Show
}

;--------
;    + ClipMenu_icon_error

ClipMenu_icon_error(ItemName, ItemPos, MyMenu) {
MsgBox ItemName "`nCheck path to see if icon files exist and correctly named.`nIf not, visit https://github.com/xypha/AHK-v2-scripts/blob/main/icons/ClipMenu/",, 262144 ; 262144 = Always-on-top
}

;--------
;    + ClipTrim

ClipTrim(number) {
Try trimmed := SubStr(StrReplace(ClipArr[number], "`r`n", A_Space), 1, 90) ; show first 90 characters, replace new line with Space
Catch as e {
    ClipArr.InsertAt(number, "[~Err0r~]: " Type(e) " - " e.Message)
    Return "[~Err0r~]"
    }
Return trimmed
}

;--------
;    + SendClipFn

SendClipFn(item, position, ClipMenu) {
PasteThis(ClipArr[position])
}

;------------------------------------------------------------------------------
;  = Paste instead of Send
; Modified from https://www.autohotkey.com/boards/viewtopic.php?p=483549#p483549 and https://www.autohotkey.com/boards/viewtopic.php?p=483588#p483588
; alternative to inbuilt command - EditPaste String, Control [, WinTitle, WinText, ExcludeTitle, ExcludeText]

;    + PasteThis

PasteThis(pasteText) {
If StrLen(pasteText) <= 15  ; 15; If short text, Send keystrokes instead of paste
    SendText pasteText      ; text mode to prevent unintended key press when text contains '^+!#{}'
Else Paste_via_clipboard(pasteText)
}

;--------
;    + Paste_via_clipboard

Paste_via_clipboard(pasteText) {
If A_Clipboard !== pasteText {
    tmp_clip := ClipboardAll()          ; preserve Clipboard
    OnClipboardChange ClipChanged, 0    ; disable callback
    A_Clipboard := pasteText            ; copy pasteText to clipboard
    tmp_clip2 := A_Clipboard
    While tmp_clip2 !== pasteText {     ; validate clipboard
        Sleep 50                        ; 50ms
        If A_Index > 5 {                ; max 250ms
            ToolTipFn(A_ThisHotkey ":: PasteThis tmp_clip copying Failed?")     ; 500ms
            OnClipboardChange ClipChanged, 1                                    ; enable callback
            Exit
            }
        }
    }
Else tmp_clip := A_Clipboard
Send "^v"                                                   ; paste
If tmp_clip !== pasteText
    SetTimer () => RestoreClip(tmp_clip, tmp_clip2), -100   ; 100ms  ; new thread - don't wait for restoration
}

;--------
;    + RestoreClip

RestoreClip(tmp_clip, tmp_clip2) {
A_Clipboard := ClipboardAll(tmp_clip)   ; restore clipboard
While tmp_clip2 == A_Clipboard {        ; validate clipboard
    Sleep 50                            ; 50ms
    If A_Index > 5 {                    ; max 250ms
        ToolTipFn(A_ThisHotkey ":: RestoreClip restoration failed!", 5000) ; 5s
        OnClipboardChange ClipChanged, 1
        Exit
        }
    }
tmp_clip := ""
tmp_clip2 := ""
OnClipboardChange ClipChanged, 1
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

; End of script code