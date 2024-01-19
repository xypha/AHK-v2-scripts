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

MyNotificationFunc("Loading AHK v2 #2 MultiClip v3", "10000", "1650", "945", "1") ; 10s

;  = Intialise ClipArr

; start clipboard array with 20 slots containing alphanumerical text
global ClipArr := ["a1","b2","c3","d4","e5","f6","g7","h8","i9","j10","k11","l12","m13","n14","o15","p16","q17","r18","s19","t20"]

; How many clipboard slots should be saved in array?
; Limit the number of slots to 20 ; change number to your needs or comment out for infinite slots
ClipArr.Capacity := 20

; run function whenever clipboard is changed
OnClipboardChange ClipChanged

; add current clipboard contents to first clipboard slot in ClipArr on start
startclip := StrReplace(A_Clipboard,"`r`n","`n")            ; Fix for SendInput sending Windows linebreaks
ClipArr.InsertAt(1, RegExReplace(startclip,"^\s+|\s+$"))    ; remove leading/trailing spaces

;  = Intialise ClipArr hotstrings

PasteVStrings(20)   ; see ClipArr Hotstrings Functions
PasteCStrings(20)

;  = Customise Tray Icon

I_Icon := A_ScriptDir "\icons\2-512.ico"
; Icon source: https://www.iconsdb.com/caribbean-blue-icons/2-icon.html     ; CC License
; I like to number scripts 1, 2, 3... and link the scripts to numpad shortcuts for easy editing
If FileExist(I_Icon)
    TraySetIcon I_Icon

;  = End auto-execute

SetTimer EndMyNotif, -1000 ; reset timer to 1s
Return

;------------------------------------------------------------------------------
; Hotkeys

;  = Check & Reload AHK

!Numpad2:: { ; CTRL & Numpad2 keys pressed together
ListLines
if WinWait(".ahk - AutoHotkey v", , 3) ; wait for listlines window to open, timeout 3s
    WinMaximize
}

^!Numpad2:: { ; CTRL & ALT & Numpad2 keys pressed together
MyNotificationFunc("Updating AHK v2 #2 MultiClip v3", "500", "1650", "985", "0") ; use sleep coz reload cancels timers
Reload
}

;------------------------------------------------------------------------------
; Hotstrings
;  = ClipArr keys

:?*x:v0+::PasteV(10) ; same as v10+ = j10

:?*x:c1+::Send "{raw}" ClipArr.Get(1) ; Send first entry in raw mode, useful when CTRL + V is disabled such as on banking sites

:?*x:c0+::PasteC(10) ; same as c10+

:?*x:c++::ClipMenuFunc(SendClipFunc) ; show ClipMenu

;------------------------------------------------------------------------------
; User-defined Functions

;  = Notification Function

MyNotificationFunc(mytext, myduration, xAxis, yAxis, timer) {
Global MyNotification
MyNotification := Gui()
MyNotification.Opt("+AlwaysOnTop -Caption +ToolWindow")  ; +ToolWindow avoids a taskbar button and an alt-tab menu item.
MyNotification.BackColor := "EEEEEE"  ; White background
MyNotification.SetFont("s9 w1000", "Arial")  ; font size 9, bold
MyNotification.Add("Text", "cBlack w181 Left", mytext)  ; black text
MyNotification.Show("x1650 y985 NoActivate")  ; NoActivate avoids deactivating the currently active window
WinMove xAxis, yAxis,,, MyNotification
if timer = 1
    SetTimer EndMyNotif, myduration * -1
if timer = 0 {
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

if DataType = 0 { ; Clipboard is now empty
    ; Tool_TipFunc("DataType: 0 - Clipboard is now empty", -1000)
    Exit
    }

if DataType = 2 { ; Clipboard contains something entirely non-text such as a picture
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

If InArr = 0
    ClipArr.InsertAt(1, Cliptemp)
else {
    ClipArr.RemoveAt(InArr)
    ClipArr.InsertAt(1, Cliptemp)
    }
}

;------------------------------------------------------------------------------
;  = ClipArr HasVal Function
; Modifed from https://www.autohotkey.com/boards/viewtopic.php?p=109173#p109173
; not for associative arrays

HasVal(haystack, needle) {
; if !(IsObject(haystack)) || (haystack.Length() = 0)
;   return -1
; optimise above code to your needs after reading lexikos' comment - https://www.autohotkey.com/boards/viewtopic.php?p=110388#p110388
for index, value in haystack
    if (value == needle) ; case-sensitive
        return index
return 0
}

;------------------------------------------------------------------------------
;  = ClipArr Hotstrings Functions

;    + PasteVStrings

PasteVStrings(number) {
loop number {
    Hotstring(":?*x:v" A_Index "+", PasteV)
    }
}

; use loop to replace similar hotstrings
; :?*:v1+::
; .
; .
; .
; :?*:v20+:: {
; PasteV(ThisHotkey)
; }

PasteV(hk) {
RegExMatch(hk, "\d+", &SubPat)
hkey := SubPat[]
if hkey = 0
    hkey := 10
try PasteThis(ClipArr.Get(hkey))
; try send ClipArr.Get(hkey) ; alternative
}

;    + PasteCStrings

PasteCStrings(number) {
loop number {
    if A_Index = 1
        continue
    Hotstring(":?*x:c" A_Index "+", PasteC)
    }
}

; :?*:c0+:: ; same as c10
; :?*:c2+::
; .
; .
; .
; :?*:c20+:: {
; PasteC(ThisHotkey)
; }


PasteC(hk) {
RegExMatch(hk, "\d+", &SubPat)
hkey := SubPat[]
if hkey = 0
    hkey := 10
PasteAll(hkey)
}

PasteAll(hkey) {
loop hkey {
    try clipVar := ClipArr.Get(A_Index)
    catch IndexError
        Result .= "`n"
    else Result .= clipVar "`n"
    }
Result := RegExReplace(Result,"^[`n]+|[`n]+$") ; remove leading/trailing LF
PasteThis(Result)
}

;------------------------------------------------------------------------------
;  = ClipArr ClipMenu Function

ClipMenuFunc(FuncName) {
global ClipMenu := Menu()
ClipMenu.Delete
ClipMenu.Add("&1  = "   ClipTrimFunc(1)   ,FuncName)
ClipMenu.Add("&2  = "   ClipTrimFunc(2)   ,FuncName)
ClipMenu.Add("&3  = "   ClipTrimFunc(3)   ,FuncName)
ClipMenu.Add("&4  = "   ClipTrimFunc(4)   ,FuncName)
ClipMenu.Add("&5  = "   ClipTrimFunc(5)   ,FuncName)
ClipMenu.Add("&6  = "   ClipTrimFunc(6)   ,FuncName)
ClipMenu.Add("&7  = "   ClipTrimFunc(7)   ,FuncName)
ClipMenu.Add("&8  = "   ClipTrimFunc(8)   ,FuncName)
ClipMenu.Add("&9  = "   ClipTrimFunc(9)   ,FuncName)
ClipMenu.Add("&0  = "   ClipTrimFunc(10)  ,FuncName)
ClipMenu.Add("&z  = "   ClipTrimFunc(11)  ,FuncName)
ClipMenu.Add("&x  = "   ClipTrimFunc(12)  ,FuncName)
ClipMenu.Add("&c  = "   ClipTrimFunc(13)  ,FuncName)
ClipMenu.Add("&v  = "   ClipTrimFunc(14)  ,FuncName)
ClipMenu.Add("&b  = "   ClipTrimFunc(15)  ,FuncName)
ClipMenu.Add("&n  = "   ClipTrimFunc(16)  ,FuncName)
ClipMenu.Add("&m = "    ClipTrimFunc(17)  ,FuncName)
ClipMenu.Add("&,    = " ClipTrimFunc(18)  ,FuncName)
ClipMenu.Add("&.    = " ClipTrimFunc(19)  ,FuncName)
ClipMenu.Add("&/   = "  ClipTrimFunc(20)  ,FuncName)
ClipMenu.Show
}

ClipTrimFunc(number) {
try ClipArr.Get(number)
catch IndexError {
    ClipArr.InsertAt(number, "")
    return ""
    }
else ClipTrim := SubStr(ClipArr.Get(number), 1, 60)
return ClipTrim
}

;    + SendClipFunc

SendClipFunc(item, position, ClipMenu) {
PasteThis(ClipArr.Get(position))
}

;------------------------------------------------------------------------------
;  = Paste instead of Send - PasteThis Function
; Modified from https://www.autohotkey.com/boards/viewtopic.php?p=483549#p483549 and https://www.autohotkey.com/boards/viewtopic.php?p=483588#p483588

PasteThis(pasteText) {
if (A_Clipboard !== pasteText) {
    OnClipboardChange ClipChanged,0
    tmp_clip := ClipboardAll()          ; preserve Clipboard
    A_Clipboard := pasteText            ; copy pastetext to clipboard
    tmp_clip2 := A_Clipboard
    While (tmp_clip2 != pasteText) {    ; validate clipboard
        Sleep 50
        if (A_Index > 5) {
            Tool_TipFunc(A_ThisHotkey ":: PasteThis Copying Failed?", -500)
            OnClipboardChange ClipChanged,1
            exit
            }
        }
    }
else tmp_clip := A_Clipboard
Send "^v"  ; paste
if (tmp_clip !== pasteText)
    SetTimer () => RestoreClip(tmp_clip, tmp_clip2), -100 ; 100ms - don't wait for restoration
}

RestoreClip(tmp_clip, tmp_clip2) {
A_Clipboard := ClipboardAll(tmp_clip)   ; restore clipboard
While (tmp_clip2 == A_Clipboard) {      ; validate clipboard
    Sleep 50
    if (A_Index > 5) {
        Tool_TipFunc(A_ThisHotkey ":: PasteThis Restoration Failed", -5000)
        OnClipboardChange ClipChanged,1
        exit
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
global ClipArr := ["a1","b2","c3","d4","e5","f6","g7","h8","i9","j10","k11","l12","m13","n14","o15","p16","q17","r18","s19","t20"]
ClipMenuFunc(SendClipFunc)  ; show menu - ClipMenu
}
