; Visit AutoHotkey (AHK) version 2 (v2) help for information - https://www.autohotkey.com/docs/v2/
; Search for below commands/functions by using control + F to search on the help webpage - https://www.autohotkey.com/docs/v2/lib/

; comments begin with semi-colon ";" at start of line or space+; " ;" in middle of line
; comments can also be show like this - "/*" comment text "*/"
; and these two methods can be combined too :)

; /* AHK 1 v2  - CONTENTS */
; Settings
; Default state of lock keys
; Auto-execute
;  = Toggle OS files
;  = Customise Tray Icon
;  = Horizontal Scrolling Group
;  = End auto-execute
; Hotkeys
;  = Check & Reload AHK
;  = Remap Keys
;  = Customise CapsLock
;  = Horizontal Scrolling
;  = Move Mouse Pointer by pixel
;  = Close or Kill an app window
;  = Adjust Window Transparency keys
;  = Monitor off
;  = Recycle Bin shortcut
;  = Add Control Panel Tools to a Menu
;  = Change the case of text
;  = Wrap Text In Quotes or Symbols keys
;  = Exchange adjacent letters
;  = Toggle Window On Top
; Capitalise first letter of a sentence
; #HotIf
;  = Firefox
;  = Telegram
;  = Windows Explorer
;    + Symbols In File Names Keys
;    + Copy full path
;    + Copy file name without path
;    + Copy file name without extension and path
; User-defined Functions
;  = Case Conversion Function
;  = HasVal Function
;  = Toggle OS Function
;  = Windows Refresh Or Run
;  = Notification Function
;  = Call Clipboard and ClipWait
;  = Wrap Text In Quotes or Symbols Function
;  = Adjust Window Transparency Function
;  = Control Panel Tools Function

;------------------------------------------------------------------------------
; Settings
; Start of script code

#Requires AutoHotkey v2.0
#SingleInstance force
#WinActivateForce
KeyHistory 500

;-------------------------------------------------------------------------------
; Default state of lock keys

SetCapsLockState "Off"
SetNumLockState "On"
SetScrollLockState "Off"

;-------------------------------------------------------------------------------
; Auto-execute
; always at the top of your script

; Show notification with parameters - text, duration in milliseconds, position on screen xAxis, yAxis, use timeout timer (1) or use sleep (0)
MyNotificationFunc("Loading AHK 1 v2", "10000", "1650", "985", "1") ; use timer for 10 seconds, position bottom right corner on 1920×1080 display resolution

;  = Toggle OS files

A_TrayMenu.Delete()                             ; delete standard menu
A_TrayMenu.Add "&Toggle OS files", ToggleOS     ; User-defined Function
A_TrayMenu.Add()                                ; add a separator
A_TrayMenu.AddStandard()                        ; restore standard menu
ToggleOSCheck()                                 ; check value of ShowSuperHidden_Status

;  = Customise Tray Icon

I_Icon := A_ScriptDir "\icons\1-512.ico"
; Icon source: https://www.iconsdb.com/caribbean-blue-icons/1-icon.html     ; CC License
; I like to number scripts 1, 2, 3... and link the scripts to numpad shortcuts for easy editing
If FileExist(I_Icon)
    TraySetIcon I_Icon

;  = Horizontal Scrolling Group

GroupAdd "HorizontalScroll1", "ahk_class ApplicationFrameWindow"       ; Modern UWP apps like calc and screen snip
GroupAdd "HorizontalScroll1", "ahk_class MozillaWindowClass"           ; Firefox
GroupAdd "HorizontalScroll1", "ahk_class SALFRAME"                     ; LibreOffice

;  = End auto-execute

SetTimer EndMyNotif, -1000 ; reset notification timer to 1s after code in auto-execute section has finished running

Return ; ends auto-execute

; Below code can be placed anywhere in your script

;------------------------------------------------------------------------------
; Hotkeys

; ^ is Control / CTRL key
; ! is ALT key
; # is Windows / Win key
; + is Shift key

;  = Check & Reload AHK

!Numpad1:: { ; CTRL & Numpad1 keys pressed together
ListLines
Sleep 500
WinMaximize "ahk - AutoHotkey"
}

^!Numpad1:: { ; CTRL & ALT & Numpad1 keys pressed together
MyNotificationFunc("Updating AHK 1 v2", "500", "1650", "985", "0") ; use sleep coz reload cancels timers
Reload
}

;-------------------------------------------------------------------------------
;  = Remap Keys

; Disable keys that I don't use or trigger accidentally too often or become annoying
; many such keys are hardware specific - desktop vs. laptop, and regional differences
; comment out the ones that don't apply to you
$ScrollLock::               ; disable Scroll Lock ; $ prefix forces keyboard hook
$NumLock::                  ; disable Num Lock
+NumpadDot::                ; Numpad delete (Modifier key - Shift)
NumpadDel::
Insert::                    ; Insert mode
+Insert::                   ; Shift + Insert
#Insert::                   ; Win + Insert
+Numpad0::                  ; Numpad Insert
NumpadIns:: {
}

!Insert::Insert    ; Use Alt + Insert to toggle the 'Insert mode' ; Source: https://gist.github.com/endolith/823381

LWin & Tab::AltTab          ; Left WIN key works as left ALT key - disables taskview

RAlt::!space       ; ALT + space brings up window menu

^RCtrl::MButton    ; press Left & Right CTRL button to simulate mouse Middle Click

RCtrl & Up::Send "{PgUp}"       ; page up
RCtrl & Down::Send "{PgDn}"     ; page down
RCtrl & Left::Send "{Home}"
RCtrl & Right::Send "{End}"

!m::WinMinimize "A"         ; Minimize active window

;-------------------------------------------------------------------------------
;  = Customise CapsLock

^CapsLock::^a        ; select all
<#CapsLock::AltTab   ; switch windows with Right Win + CapsLock

+CapsLock:: {
SetCapsLockState "On"
MyNotificationFunc("CapsLock ON", "10000", "960", "985", "1")   ; Show notification for 10s
KeyWait "Esc", "d t10" ; esc skips 10s wait and disables CapsLock immediately
SetCapsLockState "Off"
MyNotification.Destroy()
}

;-------------------------------------------------------------------------------
;  = Horizontal Scrolling
; One of these four methods should work in most situations. If not,
; search the internet for other methods and send me a msg if you find one that works for you.


; Method #1 - send window message(WM) directly to move scroll bar(SB) horizontally
; default method

+WheelUp::SendMessage 0x0114, 0, 0, ControlGetFocus("A")        ; scroll left - 0x114 is WM_HSCROLL, 0 is SB_LINELEFT
+WheelDown::SendMessage 0x0114, 1, 0, ControlGetFocus("A")      ; scroll right - 1 is SB_LINERIGHT ; same as loop 1

/* (disabled by comment)

; add additional 'Loop' command to this method to increase the speed of scrolling

+WheelUp:: {
Loop 3         ; increase the number for faster scrolling ; if number is omitted, causes infinite loop (which is BAD)
    SendMessage 0x0114, 0, 0, ControlGetFocus("A")
}

+WheelDown:: {
Loop 3
    SendMessage 0x0114, 1, 0, ControlGetFocus("A")
}

*/

;-----
; Method #2 - simulate horizontal mouse wheel action
; test if method #2 works using Win + Shift + Wheel Up/Down keys (3-key combo),
; then add window title/class to group #1 in auto-execute section
; to enable simpler Shift + Wheel Up/Down (2-key combo) via #HotIf command
; source: https://www.AutoHotkey.com/boards/viewtopic.php?t=76415

#+WheelUp::WheelLeft
#+WheelDown::WheelRight

; Group "HorizontalScroll1" is defined in auto-execute section
#HotIf WinActive("ahk_group HorizontalScroll1")                ; group #1

+WheelUp::Send "{WheelLeft}"
+WheelDown::Send "{WheelRight}"

#HotIf


/* (disabled by comment)

;-----
; Method #3 - turn on scroll lock and send arrow keys to scroll horizontally

#HotIf WinActive("ahk_group HorizontalScroll2")                ; group 2 - not yet defined in auto-execute

+WheelUp:: {
SetScrollLockState "On"
Send "{Left}"
SetScrollLockState "Off"
}

+WheelDown:: {
SetScrollLockState "On"
Send "{Right}"
SetScrollLockState "Off"
}

#HotIf

;-----
; Method #4 - send arrow keys directly if other methods don't work

#HotIf WinActive("ahk_group HorizontalScroll3")                ; group 3 - not yet defined in auto-execute

+WheelUp::Send "{Left 3}"
+WheelDown::Send "{Right 3}"

#HotIf
*/

;-------------------------------------------------------------------------------
;  = Move Mouse Pointer by pixel
; Modified from http://www.computoredge.com/AutoHotkey/Downloads/MousePrecise.ahk

#Numpad1::MouseMove -1,  1, 0, "R"    ; Win + Numpad1 move down left    ↓←
#Numpad2::MouseMove  0,  1, 0, "R"    ; Win + Numpad2 move down         ↓
#Numpad3::MouseMove  1,  1, 0, "R"    ; Win + Numpad3 move down right   ↓→
#Numpad4::MouseMove -1,  0, 0, "R"    ; Win + Numpad4 move left         ←
#Numpad5::MouseMove 960,540           ; Win + Numpad5 move center mouse • (1920×1080 display)
#Numpad6::MouseMove  1,  0, 0, "R"    ; Win + Numpad6 move right        →
#Numpad7::MouseMove -1, -1, 0, "R"    ; Win + Numpad7 move up left      ↑←
#Numpad8::MouseMove  0, -1, 0, "R"    ; Win + Numpad8 move up           ↑
#Numpad9::MouseMove  1, -1, 0, "R"    ; Win + Numpad9 move up right     ↑→

^#m::MouseMove 960,540 ; Test mouse position

;-------------------------------------------------------------------------------
;  = Close or Kill an app window
; Modified from https://superuser.com/a/1554366/391770

Alt & RButton:: { ; ALT + right mouse button ; attempt to close window
MouseGetPos ,, &id
winClass := WinGetClass("ahk_id " . id)
If (winClass != "Shell_TrayWnd")   ; exclude windows taskbar
; if (winClass != "Shell_TrayWnd" or winClass != "insert yourapp classname") ; exclude other apps using "or"
    WinClose("ahk_id " . id)  ; sends a WM_CLOSE message to the target window
    ; PostMessage 0x0112, 0xF060,,, "ahk_id " . id ; alternate method - same as pressing Alt+F4 or clicking the window's close button in its title bar:
}

; Kill window, usually unresponsive ones if WinClose fails
^!F4:: {
MouseGetPos ,, &id
WinKill ("ahk_id " . id)
}

;-------------------------------------------------------------------------------
;  = Adjust Window Transparency keys
; Modified from https://www.autohotkey.com/board/topic/667-transparent-windows/?p=148102

^+WheelUp:: { ; increases Trans value, makes the window more opaque
    Trans := GetTrans()
    if(Trans < 255)
        Trans := Trans + 20 ; add 20, change for slower/faster transition
    if(Trans >= 255)
        Trans := "Off"
    SetTrans(Trans)
}

^+WheelDown:: { ; decreases Trans value, makes the window more transparent
    Trans := GetTrans()
    if(Trans > 30)
        Trans := Trans - 20 ; subtract 20, change for slower/faster transition
    SetTrans(Trans)
}


F8:: {
SetTransMenu := Menu()
SetTransMenu.Delete
SetTransMenu.Add("&1 255 Opaque"            ,SetTransFunc)
SetTransMenu.Add("&2 190 Translucent"       ,SetTransFunc)      ; Semi-opaque
SetTransMenu.Add("&3 125 Semi-transparent"  ,SetTransFunc)
SetTransMenu.Add("&4  65 Nearly Invisible"  ,SetTransFunc)
SetTransMenu.Add("&5   0 Invisible"         ,SetTransFunc)
SetTransMenu.Show
}

;-------------------------------------------------------------------------------
;  = Recycle Bin shortcut

^del:: {
If WinActive("Recycle Bin ahk_class CabinetWClass") ; if explorer is active and recycle bin is already displayed, empty Bin
    FileRecycleEmpty
Else If Winexist("Recycle Bin ahk_class CabinetWClass") ; if explorer is inactive but showing recycle bin, activate it
    WinActivate
; Else If Winexist("ahk_class CabinetWClass") { ; if explorer is open but not showing recycle bin, change to Bin (uncomment this section if desired)
;     WinActivate
;     Sleep 1000
;     Send "{F4}"
;     Sleep 500
;     Send "{raw}::{645ff040-5081-101b-9f08-00aa002f954e}`n"
} else Run "::{645ff040-5081-101b-9f08-00aa002f954e}"    ; if explorer is not open, then open Bin in explorer
}

;-------------------------------------------------------------------------------
;  = Display Off shortcut
; modified from AHK docs

^esc:: {
Sleep 1000  ; Give user a chance to release keys (in case their release would wake up the monitor again)
SendMessage 0x0112, 0xF170, 2,, "Program Manager"  ; 0x0112 is WM_SYSCOMMAND, 0xF170 is SC_MONITORPOWER.
}

;-------------------------------------------------------------------------------
;  = Add Control Panel Tools to a Menu

#+x:: { ; Win & Shift & x
ControlPanelMenu := Menu() ; starts building a pop-up menu
ControlPanelMenu.Delete    ; deletes previously built pop-up menu, if any, and then starts adding items
ControlPanelMenu.Add("&1 Control Panel"                 ,ControlPanelFunc)
ControlPanelMenu.Add("&2 Installed Apps"                ,ControlPanelFunc)
ControlPanelMenu.Add("&3 Add/Remove Programs (Legacy)"  ,ControlPanelFunc)
ControlPanelMenu.Add("&4 Defragment Interface"          ,ControlPanelFunc)
ControlPanelMenu.Add("&5 Services"                      ,ControlPanelFunc)
ControlPanelMenu.Add("&6 Sound Mixer (Legacy)"          ,ControlPanelFunc)
ControlPanelMenu.Add("&7 Registry Editor"               ,ControlPanelFunc)
ControlPanelMenu.Add("&8 Resource Monitor"              ,ControlPanelFunc)
ControlPanelMenu.Add("&9 Windows Update"                ,ControlPanelFunc)
ControlPanelMenu.Add("&0 Windows version"               ,ControlPanelFunc) ; add more with &abc… or &symbols(/*-+)… triggers
ControlPanelMenu.Show
}

;------------------------------------------------------------------------------
;  = Change the case of text
; inspired by a v1 script from https://geekdrop.com/content/super-handy-AutoHotkey-ahk-script-to-change-the-case-of-text-in-line-or-wrap-text-in-quotes

!c:: { ; ALT + C
ChangeCaseMenu := Menu()
ChangeCaseMenu.Delete
ChangeCaseMenu.Add("&1 lower case"      ,ConvertLower)
ChangeCaseMenu.Add("&2 Sentence case"   ,ConvertSentence)
ChangeCaseMenu.Add("&3 Title Case"      ,ConvertTitle)
ChangeCaseMenu.Add("&4 UPPER CASE"      ,ConvertUpper)
ChangeCaseMenu.Add("&5 iNVERT cASE"     ,ConvertInvert)
ChangeCaseMenu.Show
}

;------------------------------------------------------------------------------
;  = Wrap Text In Quotes or Symbols keys
; Inspired by two old AHK v1 scripts from https://geekdrop.com/content/super-handy-autohotkey-ahk-script-to-change-the-case-of-text-in-line-or-wrap-text-in-quotes
; and https://www.autohotkey.com/board/topic/9805-easy-encloseenquote/?p=61995

#HotIf not WinActive("ahk_exe mpc-hc.exe") ; disable below in apps that don't use it or have conflicts

; ALT + number row
!1::EncText("`'","`'")        ; enclose in single quotation '' - ' U+0027 : APOSTROPHE
!2::EncText(Chr(34),Chr(34))  ; enclose in double quotation "" - " U+0022 : QUOTATION MARK
!3::EncText("(",")")          ; enclose in round breackets ()
!4::EncText("[","]")          ; enclose in square brackets []
!5::EncText("{","}")          ; enclose in flower brackets {}
!6::EncText(Chr(96),Chr(96))  ; enclose in accent/backtick ``
!7::EncText("%","%")          ; enclose in percent sign %%
!8::EncText("‘","’")          ; enclose in ‘’ - ‘ U+2018 LEFT & ’ U+2019 RIGHT SINGLE QUOTATION MARK {single turned comma & comma quotation mark}
!9::EncText("“","”")          ; enclose in “” - “ U+201C LEFT & ” U+201D RIGHT DOUBLE QUOTATION MARK {double turned comma & comma quotation mark}
!0::EncText("","")            ; remove above quotes

!q:: {
WrapTextMenu := Menu()
WrapTextMenu.Delete
WrapTextMenu.Add("&1   `'  Single Quotation `'"     ,WrapTextFunc) ; single quotation '' ; ordered in decreasing frequency of use; reorder as needed
WrapTextMenu.Add("&2   `" Double Quotation `""      ,WrapTextFunc) ; double quotation ""
WrapTextMenu.Add("&3   (  Round Brackets )"         ,WrapTextFunc) ; round breackets ()
WrapTextMenu.Add("&4   [  Square Brackets ]"        ,WrapTextFunc) ; square brackets []
WrapTextMenu.Add("&5   {  Flower Brackets }"        ,WrapTextFunc) ; flower brackets {}
WrapTextMenu.Add("&6   ``  Accent/Backtick ``"      ,WrapTextFunc) ; accent/backtick ``
WrapTextMenu.Add("&7  `% Percent Sign `%"           ,WrapTextFunc) ; percent sign %%
WrapTextMenu.Add("&8   ‘  Single Comma Quotation ’" ,WrapTextFunc) ; single turned comma ‘’
WrapTextMenu.Add("&9   “ Double Comma Quotation ”"  ,WrapTextFunc) ; double turned comma “”
WrapTextMenu.Add("&0  Remove all"                   ,WrapTextFunc) ; remove quotes
WrapTextMenu.Show
}

#HotIf

;------------------------------------------------------------------------------
;  = Exchange adjacent letters
; place cursor between 2 letters. The letters reverse positions - `ab|c` becomes `ac|b`.
; Modified from http://www.computoredge.com/AutoHotkey/Downloads/LetterSwap.ahk

$!l:: { ; ALT + L
clipSave := ClipboardAll()
A_Clipboard := ""
Send "{Left}+{Right 2}^c"
CallClipboardShort(2) ; 2s
SwappedLetters := SubStr(A_Clipboard,2) . SubStr(A_Clipboard,1,1)
Send SwappedLetters "{Left}"
A_Clipboard := clipSave
clipSaved := ""
}

;------------------------------------------------------------------------------
;  = Toggle Window On Top
; Modified from https://www.autohotkey.com/board/topic/94627-button-for-always-on-top/?p=596509

!t:: {                          ; ALT + t
Title_When_On_Top := "! "       ; change title "! " as required
t := WinGetTitle("A")
ExStyle := WinGetExStyle(t)
if (ExStyle & 0x8) {            ; 0x8 is WS_EX_TOPMOST
    WinSetAlwaysOnTop 0, t      ; Turn OFF and remove Title_When_On_Top
    WinSetTitle (RegexReplace(t, Title_When_On_Top)), "A"
} else {
    WinSetAlwaysOnTop 1, t      ; Turn ON and add Title_When_On_Top
    WinSetTitle Title_When_On_Top . t, t
    }
}

;------------------------------------------------------------------------------
; Capitalise first letter of a sentence
; modified from a script by Xtra - https://www.autohotkey.com/board/topic/132938-auto-capitalize-first-letter-of-sentence/?p=719739

~NumpadEnter:: ; triggers ; add or disable one or more as needed
~Enter::
~NumpadDot::
~.::
~!::
~?::
{
cfc1 := InputHook("L1 V C","{space}{LShift}{RShift}{CapsLock}", "a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,x,y,z") ; captures 1st character, visible, case sensitive ; .a → .A
cfc1.Start()
cfc1.Wait()
if (cfc1.EndReason = "Match") {
    if (A_ThisHotkey = "~!" || A_ThisHotkey = "~?") ; if ! or ? is the trigger, then add a space b/w trigger and 1st character ; !a → ! A  and ?b → ? B
        send "{Backspace} +" cfc1.Input
    else
        send "{Backspace}+" cfc1.Input ; if dot or numdot is the trigger, don't add space, coz typing website address is problematic
    exit
}
if cfc1.EndKey = "space" { ; prevent cfc2 from firing for numbers or symbols. Example: 0.2ms is not changed to 0.2Ms
    cfc2 := InputHook("L1 V C","{space}{LShift}{RShift}{CapsLock}", "a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,x,y,z") ; captures 2nd character, visible, case sensitive ; . a → . A
    cfc2.Start()
    cfc2.Wait()
    if (cfc2.EndReason = "Match")
        send "{Backspace}+" cfc2.Input
    }
}
; several other AHK v1 auto-capitalisation scripts are good, such as the one by Xtra linked above
; and one from computoredge - http://www.computoredge.com/AutoHotkey/Downloads/AutoSentenceCap.ahk
; and many others that use different methods to achieve this goal. Try a few and see what works for you.

;-------------------------------------------------------------------------------
; #HotIf
; Tailor keyboard shortcuts, commands and functions to specific windows, apps or pre-defined groups of both

;  = Firefox

#HotIf WinActive("ahk_exe firefox.exe")

^+o:: {      ; CTRL + Shift + O to open library / bookmark manager
Send "^t"
Sleep 500
Send "^l"
Sleep 500
Send "{raw}chrome://browser/content/places/places.xhtml`n" ; `n = {enter}
}

^+q::Return     ; disable Exit shortcut

#HotIf

;------------------------------------------------------------------------------
;  = Telegram

#HotIf WinActive("ahk_exe Telegram.exe")

^q::Send "^w"     ; minimise to tray, instead of quit

#HotIf

;------------------------------------------------------------------------------
;  = Windows Explorer

#HotIf WinActive("ahk_class CabinetWClass")

F1::F2 ; disable opening help in MS edge

; Unselect - Source: https://superuser.com/questions/78891/is-there-a-keyboard-shortcut-to-unselect-in-windows-explorer
^+a::Send "{F5}"

;-------
;    + Symbols In File Names Keys

; replace \/:*?"<>| with ＼⧸ ： ✲ ？＂＜＞｜
; comment out the ones you don't desire, like \ → ＼

; :?*:\::{U+FF3C}                     ; \ → ＼ | replace U+005C REVERSE SOLIDUS : backslash            → U+FF3C FULLWIDTH REVERSE SOLIDUS   ; disabled
:?*:/::{U+29F8}                     ; / → ⧸  | replace U+002F SOLIDUS : slash, forward slash, virgule → U+29F8 BIG SOLIDUS
:?*b0::+::{bs}{U+FF1A}              ; : → ：  | replace U+003A COLON                                  → U+FF1A FULLWIDTH COLON
:?*:*::{U+2732}                     ; * → ✲ | replace U+002A ASTERISK : star                         → U+2732 OPEN CENTRE ASTERISK
:?*:?::{U+FF1F}                     ; ? → ？ | replace U+003F QUESTION MARK                          → U+FF1F FULLWIDTH QUESTION MARK
:?*:"::{U+FF02}                     ; " → ＂ | replace U+0022 QUOTATION MARK : double quote          → U+FF02 FULLWIDTH QUOTATION MARK
:?*:<::{U+FF1C}                     ; < → ＜ | replace U+003C LESS-THAN SIGN                         → U+FF1C FULLWIDTH LESS-THAN SIGN
:?*:>::{U+FF1E}                     ; > → ＞ | replace U+003E GREATER-THAN SIGN                      → U+FF1E FULLWIDTH GREATER-THAN SIGN
:?*:|::{U+FF5C}                     ; | → ｜ | replace U+007C VERTICAL LINE : vertical bar, pipe     → U+FF5C FULLWIDTH VERTICAL LINE

; :*:*::{U+}                     ; ? → ? | replace ?     → ?

;-------
;    + Horizontal Scrolling
; Modified from https://www.autohotkey.com/boards/viewtopic.php?p=466527&sid=6dc4a701e678a7b9ee1241ab0043ebd8#p466527

+WheelUp:: {
Loop 3        ; WM_HSCROLL SB_LINELEFT
    PostMessage 0x0114, 0,, "ScrollBar1"
}

+WheelDown:: {
Loop 3        ; WM_HSCROLL SB_LINERIGHT
    PostMessage 0x0114, 1,, "ScrollBar1"
}

;-------
;    + Copy full path
; Modified from https://www.autohotkey.com/boards/viewtopic.php?p=61084#p61084

^+c:: { ; CTRL + Shift + C
CallClipboard(2) ; Timeout 2s
A_Clipboard := A_Clipboard
}
D:\Rise_of_the_Devourer 19-50.5.epub

;-------
;    + Copy file name without path

!n:: { ; ALT + N
CallClipboard(2) ; Timeout 2s
A_Clipboard := A_Clipboard
files := A_Clipboard
files := RegExReplace(files, "\w:\\|\w+\\") ; remove path
A_Clipboard := files
}

;-------
;    + Copy file name without extension and path

^!n:: { ; CTRL + ALT + N
CallClipboard(2) ; Timeout 2s
A_Clipboard := A_Clipboard
files := A_Clipboard
files := RegExReplace(files, "\w:\\|\w+\\") ; remove path
files := RegExReplace(files, "\.[\w]+(`r`n)","`n") ; remove ext, CR
files := RegExReplace(files, "\.[\w]+$") ; remove last ext
A_Clipboard := files
}
#HotIf

;-------------------------------------------------------------------------------
; User-defined Functions

;  = Case Conversion Function

ConvertLower(*) {
CallClipboard(2)
A_Clipboard := StrLower(A_Clipboard)
CaseConvert()
}

ConvertSentence(*) {
CallClipboard(2)
lowered := StrLower(A_Clipboard)
A_Clipboard := RegExReplace(lowered, "(((^\s*|([.!?]+\s*))[a-z])|\Wi\W)", "$U1") ; Code Credit #1
CaseConvert()
}

ConvertTitle(*) {
CallClipboard(2)
A_Clipboard := StrTitle(A_Clipboard)
CaseConvert()
}

ConvertUpper(*) {
CallClipboard(2)
A_Clipboard := StrUpper(A_Clipboard)
CaseConvert()
}

ConvertInvert(*) {
CallClipboard(2)
inverted := ""
Loop Parse A_Clipboard {     ; Code Credit #2
  if (StrLower(A_LoopField) == A_LoopField) ; * Code Credit #3
    inverted .= StrUpper(A_LoopField)       ; *
  else                                      ; *
    inverted .= StrLower(A_LoopField)       ; *
    }
A_Clipboard := inverted
CaseConvert()
}
; TestString      := "abcdefghijklmnopqrstuvwxyzéâäàåçêëèïîìæôöòûùÿáíóúñ`n"
;                  . "ABCDEFGHIJKLMNOPQRSTUVWXYZÉÂÄÀÅÇÊËÈÏÎÌÆÔÖÒÛÙŸÁÍÓÚÑ"
; TestInverted    := "ABCDEFGHIJKLMNOPQRSTUVWXYZÉÂÄÀÅÇÊËÈÏÎÌÆÔÖÒÛÙŸÁÍÓÚÑ`n"
;                  . "abcdefghijklmnopqrstuvwxyzéâäàåçêëèïîìæôöòûùÿáíóúñ"

CaseConvert() {
LenTemp := StrReplace(A_Clipboard, "`r`n", "`n") ; correct count for len
Len := "+{left " Strlen(LenTemp) "}"
Send "^v" ; Pastes new text
Send Len  ; and selects it
}

; Code Credit #1 NeedleRegEx pattern modified from pattern posted by ManaUser - https://www.autohotkey.com/board/topic/24431-convert-text-uppercase-lowercase-capitalized-or-inverted/?p=158295
; Code Credit #2 idea for loop from kon's post - https://www.autohotkey.com/boards/viewtopic.php?p=58417#p58417
; Code Credit #3 - 4 lines of code with a comment "; *" were adapted from a (inaccurate) answer generated from a auto-query to DuckDuckGPT by KudoAI via https://greasyfork.org/en/scripts/459849-duckduckgpt

;-------------------------------------------------------------------------------
;  = HasVal Function
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

;-------------------------------------------------------------------------------
;  = Toggle OS Function
; inspiration from a post by gonzax - https://www.autohotkey.com/board/topic/82603-toggle-hidden-files-system-files-and-file-extensions/?p=670182

ToggleOS(*) {
ToggleOSCheck()
If (ShowSuperHidden_Status = 0) { ; enable if disabled
    RegWrite "1", "REG_DWORD", "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSuperHidden"
    CheckRegWrite(ShowSuperHidden_Status)
    ToggleOSCheck()
    WindowsRefreshOrRun()
    } Else { ; disable if enabled
    RegWrite "0", "REG_DWORD", "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSuperHidden"
    CheckRegWrite(ShowSuperHidden_Status)
    ToggleOSCheck()
    WindowsRefreshOrRun()
    }
}

CheckRegWrite(key) { ; check if RegWrite was success
    if key = RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSuperHidden")
        msgbox "ToggleOS Failed"
}

ToggleOSCheck() { ; tray tick mark
Global ShowSuperHidden_Status
ShowSuperHidden_Status := RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSuperHidden")
If (ShowSuperHidden_Status = 0)
    A_TrayMenu.UnCheck "&Toggle OS files"
Else {
    ShowSuperHidden_Status := 1
    A_TrayMenu.Check "&Toggle OS files"
    }
}

;  = Windows Refresh Or Run

WindowsRefreshOrRun() {
if WinExist("ahk_class CabinetWClass") { ; refresh explorer if window exists
    WinActivate
    Sleep 500  ; change as per your system performance
    Send "{F5}" ; refresh
} else { ; open new explorer window if one doesn't already exist ; remove this section if not desired
    Run 'explorer.exe',,"Max"
    WinWait("ahk_class CabinetWClass",, 10) ; timeout 10 secs
    WinActivate
    }
}

;-------------------------------------------------------------------------------
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
    MyNotification.Destroy()
}

;-------------------------------------------------------------------------------
;  = Call Clipboard and ClipWait

CallClipboard(secs) {
clipSave := ClipboardAll()
A_Clipboard := ""
Send "^c"
If !ClipWait(secs) {
    MyNotificationFunc(A_ThisHotkey ":: Clip Failed", "2000", "1650", "985", "1") ; personal preferrence coz tooltip conflict
    ; ToolTip A_ThisHotkey ":: Clip Failed"           ; Alternatively to MyNotification
    ; SetTimer () => ToolTip(), -2000
    A_Clipboard := clipSave
    clipSave := ""
    Exit
    }
else
    Return clipSave
}

CallClipboardShort(secs) {
If !ClipWait(secs) {
    MyNotificationFunc(A_ThisHotkey ":: Clip Failed", "2000", "1650", "985", "1") ; personal preferrence coz tooltip conflict
    ; ToolTip A_ThisHotkey ":: Clip Failed"           ; Alternatively to MyNotification
    ; SetTimer () => ToolTip(), -2000
    Exit
    }
}

;------------------------------------------------------------------------------
;  = Wrap Text In Quotes or Symbols Function

WrapTextFunc(item, position, WrapTextMenu) {
If position = 1
    EncText("'","'")           ; enclose in single quotation '' - ' U+0027 : APOSTROPHE
Else if position = 2
    EncText(Chr(34),Chr(34))   ; enclose in double quotation "" - " U+0022 : QUOTATION MARK
Else if position = 3
    EncText("(",")")           ; enclose in round breackets ()
Else if position = 4
    EncText("[","]")           ; enclose in square brackets []
Else if position = 5
    EncText("{","}")           ; enclose in flower brackets {}
Else if position = 6
    EncText(Chr(96),Chr(96))   ; enclose in accent/backtick ``
Else if position = 7
    EncText("%","%")           ; enclose in percent sign %%
Else if position = 8
    EncText("‘","’")           ; enclose in ‘’ - ‘ U+2018 LEFT & ’ U+2019 RIGHT SINGLE QUOTATION MARK {single turned comma & comma quotation mark}
Else if position = 9
    EncText("“","”")           ; enclose in “” - “ U+201C LEFT & ” U+201D RIGHT DOUBLE QUOTATION MARK {double turned comma & comma quotation mark}
Else if position = 10
    EncText("","")             ; remove above quotes
}

EncText(q,p) {
CallClipboard(2) ; 2s
TextStringInitial := A_Clipboard
TextString := A_Clipboard
TextString := StrReplace(TextString, "`r`n", "`n")      ; fix carriage return + line feed for Strlen
TextString := RegExReplace(TextString,'^\s+|\s+$')      ; RegEx remove leading/trailing space
TextString := RegExReplace(TextString,'^[\[`'\(\{%`"“‘]+|^``')     ;"; remove leading  ['({%"“‘`
TextString := RegExReplace(TextString,'[\]`'\)\}%`"”’]+$|``$')     ;"; remove trailing ]')}%"”’`
TextString := q TextString p
TextString := StrReplace(TextString, "`n" p, p)
Len1 := Strlen(TextString)

; if you regularly include leading/trailing spaces within quotes, comment out above RegEx and below if statements
If (RegExMatch(TextStringInitial, "^\s+")) {   ; if initial string has Leading space
    TextString := " " TextString               ; add Leading space to string
    Len1++                                     ; add 1 to len
    }
If (RegExMatch(TextStringInitial, "\s+$")) {   ; if initial string has Trailing space
    TextString .= " "                          ; append trailing space to string
    Len1++                                     ; add 1 to len
    }

Len2 := "+{left " Len1 "}"

; Send "{raw}" TextString    ; send string with quotes
A_Clipboard := TextString    ; paste from clipboard is faster than send raw, especially for long strings
Send "^v"
Send Len2          ; and select textstring
; A_Clipboard := TextStringInitial  ; restore original text string to clipboard if desired
}

;-------------------------------------------------------------------------------
;  = Adjust Window Transparency Function

GetTrans() {
    ToolTip() ; disable previous tooltip if any
    MouseGetPos ,, &WinID
    Trans := WinGetTransparent("ahk_id " WinID)
    if(!Trans)
        Trans := 255
    return Trans
}

SetTrans(Transparency) {
    Trans := Transparency
    ToolTip("Transparency: " Trans)
    SetTimer () => ToolTip(), -500
    MouseGetPos ,, &WinID
    WinSetTransparent Trans, "ahk_id " WinID
}

SetTransFunc(item, position, SetTransMenu) {
MouseGetPos ,, &WinID
Trans := Trim(SubStr(item, 4, 3))
WinSetTransparent Trans, "ahk_id " WinID
}

;-------------------------------------------------------------------------------
;  = Control Panel Tools Function

ControlPanelFunc(item, position, ControlPanelMenu) {
if position = 1
    ComObject("shell.application").ControlPanelItem("control")      ; Control Panel
if position = 2
    run 'explorer.exe "ms-settings:appsfeatures"'                   ; Installed Apps ; Modern Add/Remove Programs
if position = 3
    ComObject("shell.application").ControlPanelItem("appwiz.cpl")   ; Add/Remove Programs ; Legacy Control Panel
if position = 4
    ComObject("shell.application").ControlPanelItem("dfrgui")       ; Defragment Interface
if position = 5
    ComObject("shell.application").ControlPanelItem("services.msc") ; Services
if position = 6
    ComObject("shell.application").ControlPanelItem("sndvol")       ; Sound Mixer (Legacy)
if position = 7
    ComObject("shell.application").ControlPanelItem("regedit")      ; Registry Editor
if position = 8
    ComObject("shell.application").ControlPanelItem("resmon.exe")   ; Resource Monitor
if position = 9
    run 'explorer.exe "ms-settings:windowsupdate"'                  ; Windows Update
if position = 10
    ComObject("shell.application").ControlPanelItem("winver")       ; Windows version
}

/*

'Short' list of commands (several personal modifications over the years - NOT comprehensive, at all)
Original source - https://www.autohotkey.com/boards/viewtopic.php?p=24584#p24584

; Already in Win + X menu
ComObject("shell.application").ControlPanelItem("compmgmt.msc")    ; #x   | Computer Management
ComObject("shell.application").ControlPanelItem("devmgmt.msc")     ; #x   | Device Manager
ComObject("shell.application").ControlPanelItem("hdwwiz.cpl")      ; #x   | Device Manager ; alt
ComObject("shell.application").ControlPanelItem("diskmgmt.msc")    ; #x   | Disk Management
ComObject("shell.application").ControlPanelItem("eventvwr.msc")    ; #x   | Event Viewer

; Added to Control Panel Tools function
ComObject("shell.application").ControlPanelItem("control")         ; Control Panel
run 'explorer.exe "ms-settings:appsfeatures"'                      ; Installed Apps ; Modern Add/Remove Programs
ComObject("shell.application").ControlPanelItem("appwiz.cpl")      ; Add/Remove Programs ; Legacy Control Panel
ComObject("shell.application").ControlPanelItem("dfrgui")          ; Defragment Interface
ComObject("shell.application").ControlPanelItem("services.msc")    ; Services
ComObject("shell.application").ControlPanelItem("sndvol")          ; Sound Mixer (Legacy)
ComObject("shell.application").ControlPanelItem("regedit")         ; Registry Editor
ComObject("shell.application").ControlPanelItem("resmon.exe")      ; Resource Monitor
run 'explorer.exe "ms-settings:windowsupdate"'                     ; Windows Update
ComObject("shell.application").ControlPanelItem("winver")          ; Windows version

; Add to Control Panel Tools as desired
Run '"::{21EC2020-3AEA-1069-A2DD-08002B30309D}"',,"Max"            ; Control Panel ; alternate
Run 'rundll32 sysdm.cpl`,EditEnvironmentVariables'                 ; Environmental Variables
ComObject("shell.application").ControlPanelItem("calc")            ; Calculator
ComObject("shell.application").ControlPanelItem("notepad")         ; Notepad
ComObject("shell.application").ControlPanelItem("snippingtool")    ; Snipping Tool ; Opens modern app
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

Others

Find more shortcuts to various sections within modern Settings app - https://winaero.com/ms-settings-commands-in-windows-10/
Shell:AppsFolder ; shortcuts to all apps in start menu

*/

; End of script code
