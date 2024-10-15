; https://github.com/xypha/AHK-v2-scripts/edit/main/No-2%20MultiClip.ahk
; Last updated 2024.10.15

; /* AHK v2 No-2 MultiClip - CONTENTS */
; Settings
; Auto-execute
;  = AHK Dark Mode
;  = Initialise ClipArr
;  = Initialise ClipArr hotstrings
;  = Customise Tray Icon
;  = End auto-execute
; Hotkeys
;  = Check & Reload AHK
; Hotstrings
;  = ClipArr keys
; User-defined functions
;  = MyNotification
;    + MyNotificationGui
;    + EndMyNotif
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
;    + RestoreClip
;  = ToolTip functions
;    + ToolTipFn
; Test
; ChangeLog

;------------------------------------------------------------------------------
; Settings

#Requires AutoHotkey v2.0
#SingleInstance force
KeyHistory 500

;------------------------------------------------------------------------------
; Auto-execute
; This section should always be at the top of your script

AHKname := "AHK v2 No-2 MultiClip v4.09"

; Show notification with parameters - text; duration in milliseconds; position on screen: xAxis, yAxis; timeout by - timer (1) or sleep (0)
MyNotificationGui("Loading " AHKname, -10000, 1550, 945, 1) ; 10000ms = 10 seconds (negative number so that timer will run only once), position bottom right corner (x-axis 1550 y-axis 985) on 1920×1080 display resolution; use timer

;--------
;  = Initialise ClipArr

Global LimitClipArr := 20
; Limit the number of slots to 20 ; customise limit to your needs
; Note:     Higher the number, higher the resource usage and slower the performance/response
; WARNING:  Removing the variable entirely may cause infinite loop / app hang

Global ClipArrFile := A_MyDocuments "\ClipArrFile.txt"
; ClipArrFile.txt is saved in default path of AHK's built-in variable: A_MyDocuments
; A_MyDocuments is the full path and name of the current user's "My Documents" folder. Usually corresponds to "C:\Users\<UserName>\Documents" (the final backslash is not included in the variable)
; txt file used instead of ini because 'values longer than 65,535 characters are likely to yield inconsistent results' when using IniRead, IniWrite commands

Global delim := "~•~"
; use a unique string of because if an array-slot contains this delimiter by accident, saving and loading array from file will cause errors
; Recommendation: 3 or more characters, preferably symbols with one or more Unicode characters that are difficult to type on standard keyboard
; ~ U+007E TILDE
; • U+2022 BULLET : black small circle

Global ClipArr := [] ; set Global variable and assign empty array

; Load array from file - inspired by https://www.autohotkey.com/boards/viewtopic.php?p=341809#p341809
If FileExist(ClipArrFile) ; check if file exists
    ClipArr := StrSplit(FileRead(ClipArrFile), delim)
Else ; load default values on start - 20 slots containing alphanumerical text
    ClipArr := ["a1","b2","c3","d4","e5","f6","g7","h8","i9","j10","k11","l12","m13","n14","o15","p16","q17","r18","s19","t20"]

; run function whenever clipboard is changed such as when Ctrl + x (Cut) or Ctrl + c (Copy) is pressed
; or when clipboard is altered by other apps/programs
OnClipboardChange ClipChanged

; add current clipboard contents to first clipboard slot in ClipArr on start
InsertInClipArr(A_Clipboard, 1) ; onstart = 1

; save `ClipArr` contents to `ClipArrFile.txt` when the script exits by any means,
; except when it is killed by something like "End Task" via Taskbar, Task Manager or similar
OnExit SaveClipArr

;--------
;  = Initialise ClipArr hotstrings

PasteVStrings(20)   ; User-defined function creates serial hotstrings
PasteCStrings(20)

;--------
;  = AHK Dark Mode
; download .ahk files from the `Lib` folder in this repo
; and save to disc at the same location as your script, inside a `Lib` folder 

#Include "%A_ScriptDir%\Lib\Dark Mode - ToolTip.ahk"    ; 2024.10.15
#Include "%A_ScriptDir%\Lib\Dark Mode - MsgBox.ahk"     ; 2024.10.15
; Dark Mode - Window Spy                                  ; 2024.10.15

;--------
;  = Customise Tray Icon

I_Icon := A_ScriptDir "\icons\2-512.ico"
; Icon source: https://www.iconsdb.com/caribbean-blue-icons/2-icon.html     ; CC License
; I like to number scripts 1, 2, 3... and link the scripts to Numpad shortcuts for easy editing -- see section on "Check & Reload AHK" below
If FileExist(I_Icon)
    TraySetIcon I_Icon

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

!Numpad2:: { ; Alt + Numpad2 keys pressed together
ListLines
If WinWait(A_ScriptFullPath " - AutoHotkey v" A_AhkVersion,, 3) ; 3s timeout ; wait for ListLines window to open
    WinMaximize
}

^!Numpad2:: { ; Ctrl + Alt + Numpad2 keys pressed together
MyNotificationGui("Updating " AHKname,,, 945, 0) ; 500ms ; use Sleep coz reload cancels timers
Reload
}

;------------------------------------------------------------------------------
; Hotstrings

;  = ClipArr keys

:?*x:v0+::PasteV(10) ; same as v10+ ; pastes value in slot #10 in ClipArr ; default value "j10"

:?*x:c1+::Send "{Raw}" ClipArr.Get(1) ; Send first entry in raw mode, useful when Ctrl + V is disabled such as on banking sites

:?*x:c0+::PasteC(10) ; same as c10+

:?*x:c++::ClipMenuFn(SendClipFn) ; show ClipMenu

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
;  = MultiClip ClipArr
; Sources - https://www.autohotkey.com/boards/viewtopic.php?p=326827#p326827
; and MultiClip v1 - https://www.autohotkey.com/boards/viewtopic.php?p=332658#p332658
; MultiClip v1 used or modified AHK v1 code from https://autohotkey.com/board/topic/4567-clipstep-step-through-multiple-clipboards-using-ctrl-x-c-v/
; and https://geekdrop.com/content/super-handy-autohotkey-ahk-script-to-change-the-case-of-text-in-line-or-wrap-text-in-quotes

;    + ClipChanged

ClipChanged(DataType) {

If DataType = 0 { ; 0 Clipboard is now empty
    ; ToolTipFn("DataType: 0 - Clipboard is now empty", -1000) ; 1s
    Exit
    }

If DataType = 2 { ; 2 Clipboard contains something entirely non-text such as a picture
    ToolTipFn("DataType: 2 - Non-text copied", -1000) ; 1s
    Exit
    }

; Else DataType = 1 ; Clipboard contains text (including files copied from Windows File Explorer)

; check and add to clipArr
InsertInClipArr(A_Clipboard)
}

;--------
;    + InsertInClipArr

InsertInClipArr(text, onStart := 0) {

Cliptemp := StrReplace(text,"`r`n","`n")        ; fix for SendInput sending Windows line-breaks

If RegExMatch(Cliptemp,"^\s+$") {               ; don't insert empty strings into clipArr
    If onStart != 1                             ; If NOT on startup/reload
        ToolTipFn("[~ Only \s ~]", -2000)       ; 2s ; show alert instead
    Exit
    }

Cliptemp := RegExReplace(Cliptemp,"^\s+|\s+$")  ; remove leading/trailing \s = [\r\n\t\f\v ]

; use Loop to check if Cliptemp is already in an array. If found, remove it and retrieve its `Index`
Loop LimitClipArr {
    If Cliptemp == ClipArr.Get(A_Index) {
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
If StrLen(ClipArr.Get(1)) > 1000 ; trim If more than 1000 characters
    ToolTipFn(SubStr(ClipArr.Get(1), 1, 1000) "`n... and more") ; 500ms
Else ToolTipFn(ClipArr.Get(1)) ; 500ms
}

;--------
;    + SaveClipArr
; use * because OnExit Callback accepts two parameters

SaveClipArr(*) {
Result := ""
Loop ClipArr.Length
    Result .= ClipArr.Get(A_Index) delim
If FileExist(ClipArrFile)       ; check if file exists
    FileRecycle ClipArrFile     ; send old file to recycle bin
    ; old clipboard contents can be retrieved by restoring ClipArrFile from recycle bin
    ; alternatively, use `FileDelete` command to delete permanently
FileAppend Result, ClipArrFile  ; create new file and save current clipArr contents
}

;--------
;    + PasteVStrings

PasteVStrings(number) {
Loop number {
    Hotstring(":?*x:v" A_Index "+", PasteV)
    }
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
Try PasteThis(ClipArr.Get(SubPat[]))
; Try Send ClipArr.Get(SubPat[]) ; alternative
}

;--------
;    + PasteCStrings

PasteCStrings(number) {
Loop number {
    If A_Index = 1  ; do not create c1+ hotstring, already assigned to "{Raw}" ClipArr.Get(1)
        Continue
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
    Try clipVar := ClipArr.Get(A_Index)
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
Global ClipMenu := Menu()
ClipMenu.Delete

; populate slots
ClipMenu.Add("&1  = "   ClipTrim(1)   ,FnName) ; Customise the shortcuts by altering the character after `&` in lines containing `ClipMenu.Add`
ClipMenu.Add("&2  = "   ClipTrim(2)   ,FnName) ; Explantation:
ClipMenu.Add("&3  = "   ClipTrim(3)   ,FnName) ; When the menu is displayed, a character preceded by an ampersand (&) can be selected by pressing the corresponding key on the keyboard.
ClipMenu.Add("&4  = "   ClipTrim(4)   ,FnName) ; To display a literal ampersand, specify two consecutive ampersands as in this example: "Save && Exit"
ClipMenu.Add("&5  = "   ClipTrim(5)   ,FnName)
ClipMenu.Add("&6  = "   ClipTrim(6)   ,FnName) ; Shortcuts correspond to the number/alphabet/symbol before `=`
ClipMenu.Add("&7  = "   ClipTrim(7)   ,FnName) ; Shortcuts are usually underlined, and consist of
ClipMenu.Add("&8  = "   ClipTrim(8)   ,FnName) ; numbers from Numpad or number row, and keys from the bottom row of QUERTY keyboard
ClipMenu.Add("&9  = "   ClipTrim(9)   ,FnName)
ClipMenu.Add("&0  = "   ClipTrim(10)  ,FnName)
ClipMenu.Add("&z  = "   ClipTrim(11)  ,FnName)
ClipMenu.Add("&x  = "   ClipTrim(12)  ,FnName)
ClipMenu.Add("&c  = "   ClipTrim(13)  ,FnName)
ClipMenu.Add("&v  = "   ClipTrim(14)  ,FnName)
ClipMenu.Add("&b  = "   ClipTrim(15)  ,FnName)
ClipMenu.Add("&n  = "   ClipTrim(16)  ,FnName)
ClipMenu.Add("&m = "    ClipTrim(17)  ,FnName) ; number of spaces between characters varies to improve readability in pop-up menu
ClipMenu.Add("&,    = " ClipTrim(18)  ,FnName) ; and can be changed to reflect your system font and display settings
ClipMenu.Add("&.    = " ClipTrim(19)  ,FnName)
ClipMenu.Add("&/   = "  ClipTrim(20)  ,FnName)
/* ; alternative method to populate slots without shortcuts and without messing about with spaces
Loop 20 {
    ClipMenu.Add(A_Index " = " ClipTrim(A_Index), FnName)
    }
*/

; show pop-up menu
ClipMenu.Show
}

;--------
;    + ClipTrim

ClipTrim(number) {
Try trimmed := SubStr(StrReplace(ClipArr.Get(number), "`n", A_Space), 1, 90) ; show first 90 characters, replace new line with Space
Catch as e {
    ClipArr.InsertAt(number, "[~Err0r~]: " Type(e) " - " e.Message)
    Return "[~Err0r~]"
    }
Return trimmed
}

;--------
;    + SendClipFn

SendClipFn(item, position, ClipMenu) {
PasteThis(ClipArr.Get(position))
}

;------------------------------------------------------------------------------
;  = Paste instead of Send
; Modified from https://www.autohotkey.com/boards/viewtopic.php?p=483549#p483549 and https://www.autohotkey.com/boards/viewtopic.php?p=483588#p483588
; alternative to inbuilt command - EditPaste String, Control [, WinTitle, WinText, ExcludeTitle, ExcludeText]

;    + PasteThis

PasteThis(pasteText) {
If StrLen(pasteText) <= 15  ; 15; If short text, Send keystrokes instead of paste
    SendText pasteText      ; text mode to prevent unintended key press when text contains '^+!#{}'
Else {
    If A_Clipboard !== pasteText {
        OnClipboardChange ClipChanged, 0    ; disable callback
        tmp_clip := ClipboardAll()          ; preserve Clipboard
        A_Clipboard := pasteText            ; copy pasteText to clipboard
        tmp_clip2 := A_Clipboard
        While tmp_clip2 !== pasteText {     ; validate clipboard
            Sleep 50 ; 50ms
            If A_Index > 5 { ; max 250ms
                ToolTipFn(A_ThisHotkey ":: PasteThis Copying Failed?") ; 500ms
                OnClipboardChange ClipChanged, 1 ; enable callback
                Exit
                }
            }
        }
    Else tmp_clip := A_Clipboard
    Send "^v"  ; paste
    If tmp_clip !== pasteText
        SetTimer () => RestoreClip(tmp_clip, tmp_clip2), -100 ; 100ms  ; new thread - don't wait for restoration
    }
}

;--------
;    + RestoreClip

RestoreClip(tmp_clip, tmp_clip2) {
A_Clipboard := ClipboardAll(tmp_clip)   ; restore clipboard
While tmp_clip2 == A_Clipboard {        ; validate clipboard
    Sleep 50 ; 50ms
    If A_Index > 5 { ; max 250ms
        ToolTipFn(A_ThisHotkey ":: PasteThis Restoration Failed", -5000) ; 5s
        OnClipboardChange ClipChanged, 1
        Exit
        }
    }
tmp_clip := ""
tmp_clip2 := ""
OnClipboardChange ClipChanged, 1
}

/* Example - PasteThis("1") - if ClipboardAll is NOT text

ClipboardAll()  2i      > 1     > 2i
A_Clipboard     2i      > 1     > 2i
pasteText       1
tmp_clip        0   > 2i             > 0
tmp_clip2       0           > 1      > 0

Example - PasteThis("1") - if ClipboardAll is SAME text

ClipboardAll()  2       > 1     > 2
A_Clipboard     2       > 1     > 2
pasteText       2
tmp_clip        0   > 2             > 0
tmp_clip2       0           > 1     > 0
*/

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
; Test

:*:test++:: {
SaveClipArr()
; save current array contents to file ; If script is reloaded after test, restore array contents by restoring file from recycle bin
A_Clipboard := "a1"
Global ClipArr := ["a1","b2","c3","d4","e5","f6","g7","h8","i9","j10","k11","l12","m13","n14","o15","p16","q17","r18","s19","t20"]
ClipMenuFn(SendClipFn)  ; show menu - ClipMenu
}

;------------------------------------------------------------------------------
; ChangeLog

/*
v4.00 - 2024.01.27
 + add variable `AHKname` to easily update script name and version in template and standalone scripts
 + add changelog
 * improve comments

v4.01 - 2024.01.27
 - remove version from file name
 + add alternative method to populate slots in ClipMenu
 - remove unnecessary variable `ClipTrim` from `ClipTrimFunc`
 + add `FileExist` command to `SaveClipArr` to prevent error on first exit

v4.02 - 2024.01.29
 * rename MyNotificationFunc to MyNotificationGui
 * improve ListLines WinWait command by using variables
 * rename Tool_TipFunc to ToolTipFunc
 - remove unnecessary variable hkey in PasteV function
 - remove unnecessary variable result in PasteAll function
 * improve RegEx in PasteAll function
 * rename ClipTrimFunc to ClipTrim
 - remove unnecessary parentheses in If commands
 + add PasteAndSend and SendAndPaste functions
 * some minor changes
 * improve comments and update headings

v4.03 - 2024.01.30
 * rename function names with `Func` in the name to `Fn` because `Func` is a class

v4.04 - 2024.02.01
 * improve comments

v4.05 - 2024.02.05
 + add defaults to 'MyNotificationGui' parameters
 - remove default values from all 'MyNotificationGui' func calls
 + add defaults to 'ToolTipFn' parameters
 - remove default values from all 'ToolTipFn' func calls
 - remove unnecessary quotation marks "" for 'MyNotificationGui' and 'ToolTipFn' parameters
 * improve 'PasteThis' - use 'Send' command if 'pasteText' is less than 16 characters
 * improve comments
 * improve changelog - use "fix" instead of "correct/update", use "+" for new additions and "-" for removals, "★" for new functions/sections instead of "*"

v4.06 - 2024.02.20
 * improve comments

v4.07 - 2024.03.15
 * change "!=" to "!==" wherever applicable to enable case sensitivity
 * change MyNotificationGui colour scheme to white text on dark background (dark mode)
 * improve clipboard change alert tooltip by removing unnecessary clipboard call and increasing character limit from 600 to 1000
 * replace < 16 with <= 15 in "PasteThis"
 * improve "ToolTipFn" by adding 'xAxis?, yAxis?' optional parameters
 * improve comments, spelling and small changes
 * change changelog order for easier access

v4.08 - 2024.10.11
 * rename file by replacing `#` with `No-` to avoid GitHub conflict with issue numbering
 ★ add `Dark ToolTip` section to adapt `ToolTipFn` function for windows dark mode
 * improve `ToolTipFn` function by removing unnecessary commands, change variable `ToolTipNo` to `WhichToolTip` (to match AHK docs) and change it from `Global` to `static` variable
 - remove `ToolTipOff` function
 * change `myduration` argument in `MyNotificationGui` function to use negative numbers because negative Sleep is smaller error than forever cycling SetTimer AND to match `ToolTipFn`; consequently switch negative multiplier from SetTimer to Sleep 
 * change `ClipMenuFn` shortcut from `p++` to `c++` - more intuitive
 * change `c++` shortcut to `c1+` - more intuitive, old legacy and also match `PasteCStrings` function
 * improve `InsertInClipArr` by removing unnecessary preliminary comparison, preventing addition of empty strings to clipArr and adding conditional ToolTip
 + move Clipchange ToolTip to new function `ClipArrToolTipFn` to improve handling
 * improve `ClipTrim` to show errors and insert error message in ClipArr
 * use text mode with `Send` in `PasteThis` function
 - remove `PasteAndSend` and `SendAndPaste` functions - redundant and ListLines confusion
 * rearrange/rename some function headings and update TOC
 * improve comments and small changes

v4.09 - 2024.10.15
 * rename `Dark ToolTip` section to `AHK Dark Mode` - to include all lib scripts pertaining to dark mode AHK v2
 * change dark mode ToolTip lib file from `ToolTipOptions.ahk` to `SystemThemeAwareToolTip.ahk`
 ★ add `Dark_MsgBox.ahk` and `Dark_WindowSpy` to lib and rename/modify for easier include and tracking
 * improve comments
*/
