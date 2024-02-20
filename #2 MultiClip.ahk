; /* AHK v2 #2 MultiClip - CONTENTS */
; Settings
; Auto-execute
;  = Intialise ClipArr
;  = Intialise ClipArr hotstrings
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
;    + PasteAndSend
;    + SendAndPaste
;  = ToolTip SetTimer
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

AHKname := "AHK v2 #2 MultiClip v4.06"

; Show notification with parameters - text; duration in milliseconds; position on screen: xAxis, yAxis; timeout by - timer (1) or sleep (0)
MyNotificationGui("Loading " AHKname, 10000, 1550, 945, 1) ; 10000ms = 10 seconds, position bottom right corner (x-axis 1550 y-axis 985) on 1920×1080 display resolution; use timer

;  = Intialise ClipArr

Global LimitClipArr := 20
; Limit the number of slots to 20 ; customise limit to your needs
; Note:     Higher the number, higher the resource usage and slower the performance/response
; WARNING:  Removing the variable may cause infinite loop / app hang

Global ClipArrFile := A_MyDocuments "\ClipArrFile.txt"
; ClipArrFile.txt is saved in default path of AHK's built-in variable: A_MyDocuments
; A_MyDocuments is the full path and name of the current user's "My Documents" folder. Usually corresponds to "C:\Users\<UserName>\Documents" (the final backslash is not included in the variable)
 
Global delim := "~•~"
; use a unique string because if an array-slot contains this delimiter by accident, saving and loading array from file will cause errors

Global ClipArr := [] ; set Global variable and assign empty array

; Load array from file - inspired by https://www.autohotkey.com/boards/viewtopic.php?p=341809#p341809
If FileExist(ClipArrFile) ; check if file exists
    ClipArr := StrSplit(FileRead(ClipArrFile), delim)
Else ; load default values on start - 20 slots containing alphanumerical text
    ClipArr := ["a1","b2","c3","d4","e5","f6","g7","h8","i9","j10","k11","l12","m13","n14","o15","p16","q17","r18","s19","t20"] 

; run function whenever clipboard is changed
; such as Ctrl + x (Cut), Ctrl + c (Copy) or by other apps/programs
OnClipboardChange ClipChanged

; add current clipboard contents to first clipboard slot in ClipArr on start
InsertInClipArr(A_Clipboard)

OnExit SaveClipArr
; save `ClipArr` contents to `ClipArrFile.txt` when the script exits by any means,
; except when it is killed by something like "End Task" via Taskbar, Task Manager or similar

;  = Intialise ClipArr hotstrings

PasteVStrings(20)   ; User-defined function
PasteCStrings(20)

;  = Customise Tray Icon

I_Icon := A_ScriptDir "\icons\2-512.ico"
; Icon source: https://www.iconsdb.com/caribbean-blue-icons/2-icon.html     ; CC License
; I like to number scripts 1, 2, 3... and link the scripts to Numpad shortcuts for easy editing -- see section on "Check & Reload AHK"
If FileExist(I_Icon)
    TraySetIcon I_Icon

;  = End auto-execute

SetTimer EndMyNotif, -1000 ; Reset notification timer to 1s after code in auto-execute section has finished running
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

!Numpad2:: { ; Ctrl + Numpad2 keys pressed together
ListLines
If WinWait(A_ScriptFullPath " - AutoHotkey v" A_AhkVersion,, 3) ; 3s timeout ; wait for ListLines window to open
    WinMaximize
}

^!Numpad2:: { ; Ctrl + Alt + Numpad2 keys pressed together
MyNotificationGui("Updating " AHKname,,, 945, 0) ; use Sleep coz reload cancels timers
Reload
}

;------------------------------------------------------------------------------
; Hotstrings

;  = ClipArr keys

:?*x:v0+::PasteV(10) ; same as v10+ = j10

:?*x:c++::Send "{Raw}" ClipArr.Get(1) ; Send first entry in raw mode, useful when Ctrl + V is disabled such as on banking sites

:?*x:c0+::PasteC(10) ; same as c10+

:?*x:p++::ClipMenuFn(SendClipFn) ; show ClipMenu

;------------------------------------------------------------------------------
; User-defined functions

;  = MyNotification
; search for `ToolTipFn` for alternative

;    + MyNotificationGui

MyNotificationGui(mytext, myduration := 500, xAxis := 1550, yAxis := 985, timer := 1) {
Global MyNotification := Gui("+AlwaysOnTop -Caption +ToolWindow")   ; +ToolWindow avoids a taskbar button and an Alt-Tab menu item.
MyNotification.BackColor := "EEEEEE"                ; White background, can be any RGB colour (it will be made transparent below)
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

;------------------------------------------------------------------------------
;  = MultiClip ClipArr
; Code sources - https://www.autohotkey.com/boards/viewtopic.php?p=326827#p326827
; and MultiClip v1 - https://www.autohotkey.com/boards/viewtopic.php?p=332658#p332658
; MultiClip v1 used or modified AHK v1 code from https://autohotkey.com/board/topic/4567-clipstep-step-through-multiple-clipboards-using-ctrl-x-c-v/
; and https://geekdrop.com/content/super-handy-autohotkey-ahk-script-to-change-the-case-of-text-in-line-or-wrap-text-in-quotes

;    + ClipChanged

ClipChanged(DataType) {

If DataType = 0 { ; Clipboard is now empty
    ; ToolTipFn("DataType: 0 - Clipboard is now empty", -1000) ; 1s
    Exit
    }

If DataType = 2 { ; Clipboard contains something entirely non-text such as a picture
    ToolTipFn("DataType: 2 - Non-text copied", -1000) ; 1s
    Exit
    }

; Else DataType = 1 ; Clipboard contains text (including files copied from Windows File Explorer)


; clipboard change alert tooltip
ToolTipFn(SubStr(A_Clipboard, 1, 600)) ; 500ms

InsertInClipArr(A_Clipboard)
}

;--------
;    + InsertInClipArr

InsertInClipArr(text) {

If text == ClipArr.Get(1)
    Return

; Cliptemp cleanup
Cliptemp := StrReplace(text,"`r`n","`n")        ; fix for SendInput sending Windows linebreaks
Cliptemp := RegExReplace(Cliptemp,"^\s+|\s+$")  ; remove leading/trailing \s = [\r\n\t\f\v ]


; if Cliptemp is already in a slot ≠ 1, then remove it
Loop LimitClipArr {
    If Cliptemp == ClipArr.Get(A_Index) {
        ClipArr.RemoveAt(A_Index)
        Break
        }
    }
ClipArr.InsertAt(1, Cliptemp)   ; insert current clipboard contents in the first slot
ClipArr.Length := LimitClipArr  ; reset number of slots to previously defined limit
}

;--------
;    + SaveClipArr

SaveClipArr(*) {
Result := ""
Loop ClipArr.Length
    Result .= ClipArr.Get(A_Index) delim
If FileExist(ClipArrFile)       ; check if file exists
    FileRecycle ClipArrFile     ; send old file to recycle bin
    ; old script contents can be retrieved by restoring ClipArrFile from recycle bin
FileAppend Result, ClipArrFile  ; create new file and save current cliparr contents
}

;--------
;    + PasteVStrings

PasteVStrings(number) {
Loop number {
    Hotstring(":?*x:v" A_Index "+", PasteV)
    }
}

/* use Loop to replace similar hotstrings
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

/* use Loop to replace similar hotstrings
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
/* ; alternative method to populate slots without shortcuts and messing around with spaces
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
Try ClipArr.Get(number)
Catch IndexError {
    ClipArr.InsertAt(number, "")
    Return ""
    } 
Return SubStr(ClipArr.Get(number), 1, 60)
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
If StrLen(pasteText) < 16 ; If short text, Send keystrokes instead of paste
    Send pasteText
Else {
    If A_Clipboard !== pasteText {
        OnClipboardChange ClipChanged,0
        tmp_clip := ClipboardAll()          ; preserve Clipboard
        A_Clipboard := pasteText            ; copy pastetext to clipboard
        tmp_clip2 := A_Clipboard
        While tmp_clip2 != pasteText {      ; validate clipboard
            Sleep 50 ; 50ms
            If A_Index > 5 { ; max 250ms
                ToolTipFn(A_ThisHotkey ":: PasteThis Copying Failed?") ; 500ms
                OnClipboardChange ClipChanged,1
                Exit
                }
            }
        }
    Else tmp_clip := A_Clipboard
    Send "^v"  ; paste
    If tmp_clip !== pasteText
        SetTimer () => RestoreClip(tmp_clip, tmp_clip2), -100 ; 100ms - don't wait for restoration
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
        OnClipboardChange ClipChanged,1
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

;--------
;    + PasteAndSend

PasteAndSend(pasteTxt, sendTxt) {
PasteThis(pasteTxt)
Send sendTxt
}

;--------
;    + SendAndPaste

SendAndPaste(sendTxt, pasteTxt) {
Send sendTxt
PasteThis(pasteTxt)
}

;------------------------------------------------------------------------------
;  = ToolTip SetTimer

;    + ToolTipFn

ToolTipFn(mytext, myduration := -500) { ; 500ms
ToolTip ; turn off any previous tooltip
ToolTip ToolText
SetTimer () => ToolTip(), ToolDuration
}

;------------------------------------------------------------------------------
; Test

:*:test++:: {
SaveClipArr
; save current array contents to file ; If script is reloaded after test, restore array contents by restoring file from recycle bin
A_Clipboard := "a1"
Global ClipArr := ["a1","b2","c3","d4","e5","f6","g7","h8","i9","j10","k11","l12","m13","n14","o15","p16","q17","r18","s19","t20"]
ClipMenuFn(SendClipFn)  ; show menu - ClipMenu
}

;------------------------------------------------------------------------------
; ChangeLog

/*
v4.06 - 2024.02.20
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

v4.04 - 2024.02.01
 * improve comments

v4.03 - 2024.01.30
 * rename function names with `Func` in the name to `Fn` because `Func` is a class

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

v4.01 - 2024.01.27
 - remove version from file name
 + add alternative method to populate slots in ClipMenu
 - remove unnecessary variable `ClipTrim` from `ClipTrimFunc`
 + add `FileExist` command to `SaveClipArr` to prevent error on first exit

v4.00 - 2024.01.27
 + add variable `AHKname` for versioning and updation of name in template and standalone scripts
 + add changelog
 * improve comments
*/
