; https://github.com/xypha/AHK-v2-scripts/edit/main/README.md
; https://github.com/xypha/AHK-v2-scripts/edit/main/AHK%20v2%20%231.ahk

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
;  = Tray Icon
; Hotkeys
;  = Check & Reload AHK
; Control Panel Tools Menu
; Capitalise first letter of a sentence
; Change the case of text
; Wrap Text In Quotes keys
; User-defined Functions
;  = Case Conversion Function
;  = HasVal Function
;  = Toggle OS Function
;  = Windows Refresh Or Run
;  = Notification Function
;  = Call Clipboard and ClipWait
;  = Wrap Text In Quotes Function
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
MyNotificationFunc("Loading AHK 1 v2", "10000", "1650", "985", "1") ; use timer for 10 seconds

;  = Toggle OS files

A_TrayMenu.Delete()                             ; delete standard menu
A_TrayMenu.Add "&Toggle OS files", ToggleOS     ; User-defined Function
A_TrayMenu.Add()                                ; add a separator
A_TrayMenu.AddStandard()                        ; restore standard menu
ToggleOSCheck()                                 ; check value of ShowSuperHidden_Status

;  = Tray Icon

I_Icon := A_ScriptDir "\icons\1-512.ico"
; Icon source: https://www.iconsdb.com/caribbean-blue-icons/1-icon.html     ; CC License
; I like to number scripts 1, 2, 3... and link the scripts to numpad shortcuts for easy editing
If FileExist(I_Icon)
    TraySetIcon I_Icon

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
; Control Panel Tools Menu

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
; Capitalise first letter of a sentence
; modified from a script posted here by Xtra - https://www.autohotkey.com/board/topic/132938-auto-capitalize-first-letter-of-sentence/?p=719739

~NumpadEnter:: ; triggers
~Enter::
~.::
~!::
~?::
{
cfc1 := InputHook("L1 V C"," {LShift}{RShift}", "a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,x,y,z") ; captures 1st character, visible, case sensitive ; .a → .A (trigger alphabet)
cfc1.Start()
cfc1.Wait()
if (cfc1.EndReason = "Match") {
    send "{Backspace}+" cfc1.Input
    exit
}
cfc2 := InputHook("L1 V C"," {LShift}{RShift}", "a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,x,y,z") ; captures 2nd character, visible, case sensitive ; . a → . A (trigger space alphabet)
cfc2.Start()
cfc2.Wait()
if (cfc2.EndReason = "Match")
    send "{Backspace}+" cfc2.Input
}

; there are several other AHK v1 auto-capitalisation scripts, and some of them are good -
; such as the one by Xtra linked above and one from computoredge - http://www.computoredge.com/AutoHotkey/Downloads/AutoSentenceCap.ahk

;------------------------------------------------------------------------------
; Change the case of text
; inspired by a v1 script from https://geekdrop.com/content/super-handy-AutoHotkey-ahk-script-to-change-the-case-of-text-in-line-or-wrap-text-in-quotes

!c:: { ; ALT + C
WrapQuotesMenu := Menu()
WrapQuotesMenu.Delete
WrapQuotesMenu.Add("&1 lower case"      ,ConvertLower)
WrapQuotesMenu.Add("&2 Sentence case"   ,ConvertSentence)
WrapQuotesMenu.Add("&3 Title Case"      ,ConvertTitle)
WrapQuotesMenu.Add("&4 UPPER CASE"      ,ConvertUpper)
WrapQuotesMenu.Add("&5 iNVERT cASE"     ,ConvertInvert)
WrapQuotesMenu.Show
}

;------------------------------------------------------------------------------
; Wrap Text In Quotes keys
; Inspired by two old AHK v1 scripts from https://geekdrop.com/content/super-handy-autohotkey-ahk-script-to-change-the-case-of-text-in-line-or-wrap-text-in-quotes
; and https://www.autohotkey.com/board/topic/9805-easy-encloseenquote/?p=61995

#HotIf not WinActive("ahk_exe mpc-hc.exe") ; disable below in apps that don't use it or have conflicts

; ALT + number row
!1::EncQuote("`'","`'")         ; enclose in single quotes '' - ' U+0027 : APOSTROPHE
!2::EncQuote(Chr(34),Chr(34))   ; enclose in double quotes "" - " U+0022 : QUOTATION MARK
!3::EncQuote("(",")")           ; enclose in round breackets ()
!4::EncQuote("[","]")           ; enclose in square brackets []
!5::EncQuote("{{}","{}}")       ; enclose in flower brackets {}
!6::EncQuote(Chr(96),Chr(96))   ; enclose in accent/backtick ``
!7::EncQuote("%","%")           ; enclose in percentage sign %%
!8::EncQuote("‘","’")           ; enclose in single quotes ‘’ - ‘ U+2018 LEFT & ’ U+2019 RIGHT SINGLE QUOTATION MARK {single turned comma & comma quotation mark}
!9::EncQuote("“","”")           ; enclose in double quotes “” - “ U+201C LEFT & ” U+201D RIGHT DOUBLE QUOTATION MARK {double turned comma & comma quotation mark}
!0::EncQuote("","")             ; remove above quotes

!q:: {
WrapQuotesMenu := Menu()
WrapQuotesMenu.Delete
WrapQuotesMenu.Add("&1  `'  Single Quotes"    ,WrapQuotesFunc) ; single quotes ''
WrapQuotesMenu.Add("&2  `"  Double Quotes"    ,WrapQuotesFunc) ;" double quotes ""
WrapQuotesMenu.Add("&3  (  Round Breackets"   ,WrapQuotesFunc) ; round breackets ()
WrapQuotesMenu.Add("&4  [  Square Brackets"   ,WrapQuotesFunc) ; square brackets []
WrapQuotesMenu.Add("&5  *  Asterix"           ,WrapQuotesFunc) ; flower brackets {}
WrapQuotesMenu.Add("&6  ``  Accent/Backtick"  ,WrapQuotesFunc) ; accent/backtick ``
WrapQuotesMenu.Add("&7  `%  Percentage Sign"  ,WrapQuotesFunc) ; percentage sign %%
WrapQuotesMenu.Add("&8  ‘’  Single Quotation" ,WrapQuotesFunc) ; single quotes ‘’
WrapQuotesMenu.Add("&9  “”  Double Quotation" ,WrapQuotesFunc) ; double quotes “”
WrapQuotesMenu.Add("&0  Remove All Quotes"    ,WrapQuotesFunc) ; remove quotes
WrapQuotesMenu.Show
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

ToggleOSCheck() {
Global ShowSuperHidden_Status
ShowSuperHidden_Status := RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSuperHidden")
If (ShowSuperHidden_Status = 0)
    A_TrayMenu.UnCheck "&Toggle OS files"
Else {
    ShowSuperHidden_Status := 1
    A_TrayMenu.Check "&Toggle OS files"
    }
}

ToggleOS(*) {
ToggleOSCheck()
If (ShowSuperHidden_Status = 0) { ; enable if disabled
    RegWrite "1", "REG_DWORD", "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSuperHidden"
    ToggleOSCheck()
    WindowsRefreshOrRun()
    } Else { ; disable if enabled
    RegWrite "0", "REG_DWORD", "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSuperHidden"
    ToggleOSCheck()
    WindowsRefreshOrRun()
    }
}

;  = Windows Refresh Or Run

WindowsRefreshOrRun() {
if WinExist("ahk_class CabinetWClass") { ; refresh explorer if window exists
    WinActivate
    Sleep 500  ; change as per your system performance
    Send "{F5}" ; refresh
} else { ; open new explorer window if one doesn't already exist ; remove this section if not desired
    Run 'explorer.exe',,"max"
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
;  = Wrap Text In Quotes Function

WrapQuotesFunc(item, position, WrapQuotesMenu) {
If position = 1
    EncQuote("'","'")           ; enclose in single quotes ''
Else if position = 2
    EncQuote(Chr(34),Chr(34))   ; enclose in double quotes ""
Else if position = 3
    EncQuote("(",")")           ; enclose in round breackets ()
Else if position = 4
    EncQuote("[","]")           ; enclose in square brackets []
Else if position = 5
    EncQuote("{{}","{}}")       ; enclose in flower brackets {}
Else if position = 6
    EncQuote(Chr(96),Chr(96))   ; enclose in accent/backtick ``
Else if position = 7
    EncQuote("%","%")           ; enclose in percentage sign %%
Else if position = 8
    EncQuote("‘","’")           ; enclose in single quotes ‘’
Else if position = 9
    EncQuote("“","”")           ; enclose in double quotes “”
Else if position = 10
    EncQuote("","")             ; remove quotes
}

EncQuote(q,p) {
CallClipboard(2) ; 2s
QuoteStringInitial := A_Clipboard
QuoteString := A_Clipboard
QuoteString := StrReplace(QuoteString, "`r`n", "`n")
QuoteString := RegExReplace(QuoteString,'[\[\]\*`'\(\)\{\}%`"“”‘’]+|``')     ;"; remove "*`'(){}%“”‘’[]
QuoteString := RegExReplace(QuoteString,'^\s+|\s+$')     ; RegEx remove leading/trailing space
QuoteString := q QuoteString p
QuoteString := StrReplace(QuoteString, "`r`n" p, p)
Len1 := "0"
Len2 := "0"
If q ~= "{" ; RegEx match
    Len1 := Strlen(QuoteString) - 4 ; to account for extra {} in q and p
Else
    Len1 := Strlen(QuoteString)
; if you regularly include leading/trailing spaces within quotes, comment out above RegEx and below if statements
If (RegExMatch(QuoteStringInitial, "^\s+")) {   ; if initial string has Leading space
    QuoteString := " " QuoteString              ; add Leading space to string
    Len1++                                      ; add 1 to len
    }
If (RegExMatch(QuoteStringInitial, "\s+$")) {   ; if initial string has Trailing space
    QuoteString .= " "                          ; append trailing space to string
    Len1++                                      ; add 1 to len
    }
Len2 := "+{left " Len1 "}"
Send QuoteString    ; send string with quotes
Send Len2           ; and select it
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

; already in Win X menu
ComObject("shell.application").ControlPanelItem("compmgmt.msc")    ; #x   | Computer Management
ComObject("shell.application").ControlPanelItem("devmgmt.msc")     ; #x   | Device Manager
ComObject("shell.application").ControlPanelItem("hdwwiz.cpl")      ; #x   | Device Manager ; alt
ComObject("shell.application").ControlPanelItem("diskmgmt.msc")    ; #x   | Disk Management
ComObject("shell.application").ControlPanelItem("eventvwr.msc")    ; #x   | Event Viewer

; Added to Control Panel Tools function
ComObject("shell.application").ControlPanelItem("calc")            ; Calculator
ComObject("shell.application").ControlPanelItem("notepad")         ; Notepad
ComObject("shell.application").ControlPanelItem("snippingtool")    ; Snipping Tool ; Opens modern app
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
