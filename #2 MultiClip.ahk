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
; User-defined Functions
;  = Notification Function
;  = ClipArr ClipChanged function
;  = ClipArr InsertInClipArr function
;  = ClipArr SaveClipArr function
;  = ClipArr Hotstrings functions
;    + PasteVStrings
;    + PasteCStrings
;  = ClipArr ClipMenu function
;    + SendClipFunc
;  = Paste instead of Send - PasteThis Function
;  = ToolTip Function
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

AHKname := "AHK v2 #2 MultiClip v4.00"

; Show notification with parameters - text; duration in milliseconds; position on screen: xAxis, yAxis; timeout by - timer (1) or sleep (0)
MyNotificationFunc("Loading " AHKname, "10000", "1550", "945", "1") ; 10000 milliseconds = 10 seconds, position bottom right corner (x-axis 1550 y-axis 985) on 1920×1080 display resolution; use timer

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

Global ClipArr := [] ; set Global variable and load empty array

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

PasteVStrings(20)   ; see ClipArr Hotstrings functions
PasteCStrings(20)

;  = Customise Tray Icon

I_Icon := A_ScriptDir "\icons\2-512.ico"
; Icon source: https://www.iconsdb.com/caribbean-blue-icons/2-icon.html     ; CC License
; I like to number scripts 1, 2, 3... and link the scripts to Numpad shortcuts for easy editing
If FileExist(I_Icon)
    TraySetIcon I_Icon

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

!Numpad2:: { ; Ctrl + Numpad2 keys pressed together
ListLines
If WinWait(".ahk - AutoHotkey v",, 3) ; wait for ListLines window to open, timeout 3s
    WinMaximize
}

^!Numpad2:: { ; Ctrl + Alt + Numpad2 keys pressed together
MyNotificationFunc("Updating " AHKname, "500", "1550", "945", "0") ; use Sleep coz reload cancels timers
Reload
}

;------------------------------------------------------------------------------
; Hotstrings

;  = ClipArr keys

:?*x:v0+::PasteV(10) ; same as v10+ = j10

:?*x:c++::Send "{Raw}" ClipArr.Get(1) ; Send first entry in raw mode, useful when Ctrl + V is disabled such as on banking sites

:?*x:c0+::PasteC(10) ; same as c10+

:?*x:p++::ClipMenuFunc(SendClipFunc) ; show ClipMenu

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

;------------------------------------------------------------------------------
;  = ClipArr ClipChanged function
; Modified from MultiClip v1 https://www.autohotkey.com/boards/viewtopic.php?p=332658#p332658
; and https://www.autohotkey.com/boards/viewtopic.php?p=326827#p326827

ClipChanged(DataType) {

If DataType = 0 { ; Clipboard is now empty
    ; Tool_TipFunc("DataType: 0 - Clipboard is now empty", -1000)
    Exit
    }

If DataType = 2 { ; Clipboard contains something entirely non-text such as a picture
    Tool_TipFunc("DataType: 2 - Non-text copied", -1000)
    Exit
    }

; Else DataType = 1 ; Clipboard contains text (including files copied from Windows File Explorer)


; clipboard change alert tooltip
Tool_TipFunc(SubStr(A_Clipboard, 1, 600), -500)

InsertInClipArr(A_Clipboard)
}

;--------
;  = ClipArr InsertInClipArr function

InsertInClipArr(text) {

If text == ClipArr.Get(1)
    Return

; Cliptemp cleanup
Cliptemp := StrReplace(text,"`r`n","`n")        ; fix for SendInput sending Windows linebreaks
Cliptemp := RegExReplace(Cliptemp,"^\s+|\s+$")  ; remove leading/trailing \s \t \r \n


; Check if Cliptemp is already in an array and retrieve its `Index` if present
Loop LimitClipArr {
    If Cliptemp == ClipArr.Get(A_Index) {
        ClipArr.RemoveAt(A_Index)
        Break
        }
    }
ClipArr.InsertAt(1, Cliptemp)   ; insert current clipboard contents in first slot
ClipArr.Length := LimitClipArr  ; reset number of slots to previously defined limit
}

;--------
;  = ClipArr SaveClipArr function

SaveClipArr(*) {
Result := ""
Loop ClipArr.Length
    Result .= ClipArr.Get(A_Index) delim
FileRecycle ClipArrFile         ; send old file to recycle bin
FileAppend Result, ClipArrFile  ; create new file and save current cliparr contents
}

;------------------------------------------------------------------------------
;  = ClipArr Hotstrings functions

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
hkey := SubPat[]
Try PasteThis(ClipArr.Get(hkey))
; Try Send ClipArr.Get(hkey) ; alternative
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
Result := RegExReplace(Result,"^[`n]+|[`n]+$") ; remove leading/trailing LF
PasteThis(Result)
}

;------------------------------------------------------------------------------
;  = ClipArr ClipMenu function

ClipMenuFunc(FuncName) {
Global ClipMenu := Menu()
ClipMenu.Delete
ClipMenu.Add("&1  = "   ClipTrimFunc(1)   ,FuncName) ; Customise the shortcuts by altering the character after `&` in lines containing `ClipMenu.Add`
ClipMenu.Add("&2  = "   ClipTrimFunc(2)   ,FuncName) ; Explantation: 
ClipMenu.Add("&3  = "   ClipTrimFunc(3)   ,FuncName) ; When the menu is displayed, a character preceded by an ampersand (&) can be selected by pressing the corresponding key on the keyboard.
ClipMenu.Add("&4  = "   ClipTrimFunc(4)   ,FuncName) ; To display a literal ampersand, specify two consecutive ampersands as in this example: "Save && Exit"
ClipMenu.Add("&5  = "   ClipTrimFunc(5)   ,FuncName)
ClipMenu.Add("&6  = "   ClipTrimFunc(6)   ,FuncName) ; Shortcuts correspond to the number/alpabet/symbol prior to `=`
ClipMenu.Add("&7  = "   ClipTrimFunc(7)   ,FuncName) ; Shortcuts are usually underlined, and consist of
ClipMenu.Add("&8  = "   ClipTrimFunc(8)   ,FuncName) ; numbers from NumPad or number row, and keys from the bottom row of QUERTY keyboard
ClipMenu.Add("&9  = "   ClipTrimFunc(9)   ,FuncName)
ClipMenu.Add("&0  = "   ClipTrimFunc(10)  ,FuncName)
ClipMenu.Add("&z  = "   ClipTrimFunc(11)  ,FuncName)
ClipMenu.Add("&x  = "   ClipTrimFunc(12)  ,FuncName)
ClipMenu.Add("&c  = "   ClipTrimFunc(13)  ,FuncName)
ClipMenu.Add("&v  = "   ClipTrimFunc(14)  ,FuncName)
ClipMenu.Add("&b  = "   ClipTrimFunc(15)  ,FuncName)
ClipMenu.Add("&n  = "   ClipTrimFunc(16)  ,FuncName)
ClipMenu.Add("&m = "    ClipTrimFunc(17)  ,FuncName) ; number of spaces between characters vary in order to improve readability in pop-up menu
ClipMenu.Add("&,    = " ClipTrimFunc(18)  ,FuncName) ; and can be changed to reflect your system font and display settings
ClipMenu.Add("&.    = " ClipTrimFunc(19)  ,FuncName)
ClipMenu.Add("&/   = "  ClipTrimFunc(20)  ,FuncName)
ClipMenu.Show
}

ClipTrimFunc(number) {
Try ClipArr.Get(number)
Catch IndexError {
    ClipArr.InsertAt(number, "")
    Return ""
    }
Else ClipTrim := SubStr(ClipArr.Get(number), 1, 60)
Return ClipTrim
}

;--------
;    + SendClipFunc

SendClipFunc(item, position, ClipMenu) {
PasteThis(ClipArr.Get(position))
}

;------------------------------------------------------------------------------
;  = Paste instead of Send - PasteThis Function
; Modified from https://www.autohotkey.com/boards/viewtopic.php?p=483549#p483549 and https://www.autohotkey.com/boards/viewtopic.php?p=483588#p483588
; alternative to inbuilt command - EditPaste String, Control [, WinTitle, WinText, ExcludeTitle, ExcludeText]

PasteThis(pasteText) {
If (A_Clipboard !== pasteText) {
    OnClipboardChange ClipChanged,0
    tmp_clip := ClipboardAll()          ; preserve Clipboard
    A_Clipboard := pasteText            ; copy pastetext to clipboard
    tmp_clip2 := A_Clipboard
    While (tmp_clip2 != pasteText) {    ; validate clipboard
        Sleep 50
        If (A_Index > 5) {
            Tool_TipFunc(A_ThisHotkey ":: PasteThis Copying Failed?", -500)
            OnClipboardChange ClipChanged,1
            Exit
            }
        }
    }
Else tmp_clip := A_Clipboard
Send "^v"  ; paste
If (tmp_clip !== pasteText)
    SetTimer () => RestoreClip(tmp_clip, tmp_clip2), -100 ; 100ms - don't wait for restoration
}

RestoreClip(tmp_clip, tmp_clip2) {
A_Clipboard := ClipboardAll(tmp_clip)   ; restore clipboard
While (tmp_clip2 == A_Clipboard) {      ; validate clipboard
    Sleep 50
    If (A_Index > 5) {
        Tool_TipFunc(A_ThisHotkey ":: PasteThis Restoration Failed", -5000)
        OnClipboardChange ClipChanged,1
        Exit
        }
    }
tmp_clip := "", tmp_clip2 := ""
OnClipboardChange ClipChanged,1
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
;  = ToolTip Function

Tool_TipFunc(ToolText, ToolDuration) {
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
ClipMenuFunc(SendClipFunc)  ; show menu - ClipMenu
}

;------------------------------------------------------------------------------
; ChangeLog

/*

v4.00 - 2024.01.27
 * add variable `AHKname` for versioning and updation of name in template and standalone scripts
 * add changelog
 * improve comments

*/
