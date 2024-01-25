; /* AHK v2 #2 MultiClip v3 - CONTENTS */
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
;  = ClipArr ClipChanged Function
;  = ClipArr HasVal Function
;  = ClipArr Hotstrings Functions
;    + PasteVStrings
;    + PasteCStrings
;  = ClipArr ClipMenu Function
;    + SendClipFunc
;  = Paste instead of Send - PasteThis Function
;  = ToolTip Function
; Test

;------------------------------------------------------------------------------
; Settings

#Requires AutoHotkey v2.0
#SingleInstance force
KeyHistory 500

;------------------------------------------------------------------------------
; Auto-execute

MyNotificationFunc("Loading AHK v2 #2 MultiClip v3", "10000", "1550", "945", "1") ; 10s

;  = Intialise ClipArr

; Start clipboard array with 20 slots containing alphanumerical text
Global ClipArr := ["a1","b2","c3","d4","e5","f6","g7","h8","i9","j10","k11","l12","m13","n14","o15","p16","q17","r18","s19","t20"]

; Limit the number of slots to 20 ; change number to your needs or comment out for infinite slots
ClipArr.Capacity := 20

; run function whenever clipboard is changed
OnClipboardChange ClipChanged

; add current clipboard contents to first clipboard slot in ClipArr on start
startClip := StrReplace(A_Clipboard,"`r`n","`n")            ; Fix for SendInput sending Windows linebreaks
ClipArr.InsertAt(1, RegExReplace(startClip,"^\s+|\s+$"))    ; Remove leading/trailing spaces

;  = Intialise ClipArr hotstrings

PasteVStrings(20)   ; see ClipArr Hotstrings Functions
PasteCStrings(20)

;  = Customise Tray Icon

I_Icon := A_ScriptDir "\icons\2-512.ico"
; Icon source: https://www.iconsdb.com/caribbean-blue-icons/2-icon.html     ; CC License
; I like to number scripts 1, 2, 3... and link the scripts to Numpad shortcuts for easy editing
If FileExist(I_Icon)
    TraySetIcon I_Icon

;  = End auto-execute

SetTimer EndMyNotif, -1000 ; reset timer to 1s
Return

;------------------------------------------------------------------------------
; Hotkeys

;  = Check & Reload AHK

!Numpad2:: { ; Ctrl + Numpad2 keys pressed together
ListLines
If WinWait(".ahk - AutoHotkey v",, 3) ; wait for ListLines window to open, timeout 3s
    WinMaximize
}

^!Numpad2:: { ; Ctrl + Alt + Numpad2 keys pressed together
MyNotificationFunc("Updating AHK v2 #2 MultiClip v3", "500", "1550", "945", "0") ; use sleep coz reload cancels timers
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
;  = ClipArr ClipChanged Function
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

; DataType = 1 ; Clipboard contains text (including files copied from an Explorer)

; Cliptemp cleanup
Cliptemp := StrReplace(A_Clipboard,"`r`n","`n") ; Fix for SendInput sending Windows linebreaks
Cliptemp := RegExReplace(Cliptemp,"^\s+|\s+$")  ; remove leading/trailing spaces

; clipboard change alert tooltip
Tool_TipFunc(SubStr(Cliptemp, 1, 600), -500)

; Check if Cliptemp is already in an array and retrieve its `Index` if present
InArr := HasVal(ClipArr, Cliptemp)

If InArr !== 0
    ClipArr.RemoveAt(InArr)
ClipArr.InsertAt(1, Cliptemp)
}

;------------------------------------------------------------------------------
;  = ClipArr HasVal Function
; Modifed from https://www.autohotkey.com/boards/viewtopic.php?p=109173#p109173
; not for associative arrays

HasVal(haystack, needle) {
/* ; optimise this code to your needs after reading lexikos' comment - https://www.autohotkey.com/boards/viewtopic.php?p=110388#p110388
If !(IsObject(haystack)) || (haystack.Length() = 0)
  Return -1
*/
For index, value in haystack
    If (value == needle) ; case-sensitive
        Return index
Return 0
}

;------------------------------------------------------------------------------
;  = ClipArr Hotstrings Functions

;    + PasteVStrings

PasteVStrings(number) {
Loop number {
    Hotstring(":?*x:v" A_Index "+", PasteV)
    }
}

/* use loop to replace similar hotstrings
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

;    + PasteCStrings

PasteCStrings(number) {
Loop number {
    If A_Index = 1  ; do not create c1+ hotstring, already assigned to "{Raw}" ClipArr.Get(1) 
        Continue
    Hotstring(":?*x:c" A_Index "+", PasteC)
    }
}

/* use loop to replace similar hotstrings
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
;  = ClipArr ClipMenu Function

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
A_Clipboard := "a1"
Global ClipArr := ["a1","b2","c3","d4","e5","f6","g7","h8","i9","j10","k11","l12","m13","n14","o15","p16","q17","r18","s19","t20"]
ClipMenuFunc(SendClipFunc)  ; show menu - ClipMenu
}
