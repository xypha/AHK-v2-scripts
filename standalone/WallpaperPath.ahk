; https://github.com/xypha/AHK-v2-scripts/edit/main/standalone/WallpaperPath.ahk
; Last updated 2024.11.20

    /* CONTENTS v6.02 */
; Hotkey
; User-defined functions
;  = Clipboard function
;    + CallClipWait
;  = ToolTip function
;    + ToolTipFn
;  = Extract wallpaper location from registry (pure AHK)
;    + WallpaperPath_v2
;    + WallpaperPath_v3
;    + WallpaperPath_v4
;  = Extract wallpaper location from registry (AHK + PowerShell)
;    + WallpaperPath_v5() - Output through clipboard
;    + WallpaperPath_v6() - Output through array without altering clipboard
;  = Run PowerShell commands through AHK
;    + RunPS(commands)
;  = Locate TranscodedWallpaper
; ChangeLog

#Requires AutoHotkey v2.0
#SingleInstance force

;------------------------------------------------------------------------------
; Hotkey

; #HotIf WinActive("ahk_class WorkerW ahk_exe explorer.exe")
; commented out - enable shortcut on desktop only

#W::{ ; Win + W
MsgBox "WallpaperPath_v4`n" WallpaperPath_v4()
}

; #HotIf

;------------------------------------------------------------------------------
; User-defined functions

;  = Clipboard function

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
;  = Extract wallpaper location from registry (pure AHK)
; below code was modified from v1 - https://gist.github.com/raveren/bac5196d2063665d2154#file-aio-ahk-L741
; This source has code for multi-monitor setups that is not included here.
; Look for `openWallpaperUnderMouse()` and `getMonitorUnderMouse()` functions over there if you need it.

/*  ; Example input (refBinary) from registry:
7AC30100FEB72F00001E0000E01000009D65B985BF76D70144003A005C00570061006C006C007000610070006500720073005C004C004900460045007300740079006C0065005F004E006500770073005F004D006900580074007500720065005F0049006D0061006700650073005F0031003400360037005F005F003000360037002E006A00700067000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000

  ; Output - return from `WallpaperPath_v4()` -
D:\Wallpapers\LIFEstyle_News_MiXture_Images_1467__067.jpg
*/

;--------
;    + WallpaperPath_v2

WallpaperPath_v2() {

; Read the binary value from the registry
regBinary := RegRead("HKEY_CURRENT_USER\Control Panel\Desktop", "TranscodedImageCache")

; remove first 48 characters
regBinary := SubStr(regBinary, 49)

; format "F5A6" => "F5,A6"
Loop Parse regBinary
    regContents .= A_LoopField (Mod(A_Index, 2) ? "" : ",")

ascii := '!"#$%&`'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_``abcdefghijklmnopqrstuvwxyz{|}~' ;" comment ; escape `' (apostrophe) `` (accent)
consequtiveZeroes := 0
ConvString := ""
isUnicode := 0
Loop Parse, regContents, "," {

    ; detect string end \0
    If (consequtiveZeroes > 1)
        Break

    ; If previous char is in Unicode
    If isUnicode = 1 {
        char := Chr("0x" A_LoopField char)
        ConvString .= char
        isUnicode := 0
        Continue
        }

    ; convert HEX to ASCII
    if (A_LoopField = 0)
        consequtiveZeroes++
    else {
        char := Chr("0x" A_LoopField)

        ; If ASCII character, add to path string
        If InStr(ascii, char) {
            isUnicode := 0
            ConvString .= char
            }

        ; If not ASCII, then it is Unicode
        Else {
            char := A_LoopField
            isUnicode := 1
            }
        consequtiveZeroes := 0
        }
    }
Return ConvString
}

;--------
;    + WallpaperPath_v3

WallpaperPath_v3() {

; Read the binary value from the registry
regBinary := RegRead("HKEY_CURRENT_USER\Control Panel\Desktop", "TranscodedImageCache")

; remove first 48 characters
regBinary := SubStr(regBinary, 49)


; format "F5A6" => "F5,A6"
Loop Parse regBinary
    regContents .= A_LoopField (Mod(A_Index, 2) ? "" : ",")

regArr := StrSplit(regContents, ",", A_Space)

ConvString := ""
Loop regArr.Length {
    If not Mod(A_Index, 2)
        Continue
    Else char := Chr("0x" regArr[A_Index + 1] regArr[A_Index])

    If char = "" ; OR char = 0
        Break
    Else ConvString .= char
    }
Return ConvString
}

;--------
;    + WallpaperPath_v4

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

;------------------------------------------------------------------------------
;  = Extract wallpaper location from registry (AHK + PowerShell)

;    + WallpaperPath_v5() - Output through clipboard
; Note: A brief blue window will be seen for a few seconds. Wait for it to disappear automatically. This behaviour is expected, because this window is executing the PowerShell .ps1 file.

WallpaperPath_v5() {
A_Clipboard := ""
ps1_path := A_MyDocuments "\WallpaperPath.ps1"
If not FileExist(ps1_path) {
    ps_commands := "    ; continuation section
(                       ; Contents of .ps1 file - modified from https://superuser.com/a/1386340/391770
$TIC = (Get-ItemProperty 'HKCU:\Control Panel\Desktop' TranscodedImageCache -ErrorAction Stop).TranscodedImageCache
$result = [System.Text.Encoding]::Unicode.GetString($TIC) -replace '(.+)([A-Z]:[0-9a-zA-Z\\])+','$2'
Set-Clipboard -Value $result
)"
    FileAppend ps_commands, ps1_path, "`n UTF-8"
    }

Run('powershell.exe -windowstyle hidden -ExecutionPolicy Bypass -file "' ps1_path '"')
Return CallClipWait(2, 1)
}

;--------
;    + WallpaperPath_v6() - Output through array without altering clipboard
; Note: A brief black window will be seen for a few seconds. Wait for it to disappear automatically. This behaviour is expected, because this window is executing the PowerShell commands.

WallpaperPath_v6() {
ps_commands := " ; continuation section
(
$TIC = (Get-ItemProperty 'HKCU:\Control Panel\Desktop' TranscodedImageCache -ErrorAction Stop).TranscodedImageCache
[System.Text.Encoding]::Unicode.GetString($TIC) -replace '(.+)([A-Z]:[0-9a-zA-Z\\])+','$2'
)"

resultArr := StrSplit(RunPS(ps_commands), "`n", "`r")

Return resultArr[resultArr.Length] ; show last line only
}

;------------------------------------------------------------------------------
;  = Run PowerShell commands through AHK

;    + RunPS(commands)
; modified from AutoHotkey.chm::/docs/lib/Run.htm#ExecScript and https://www.autohotkey.com/boards/viewtopic.php?p=123912#p123912
; for more versatile implementation of this type of function, visit https://www.autohotkey.com/boards/viewtopic.php?f=83&t=133668

RunPS(commands) {
shell := ComObject("WScript.Shell")
exec := shell.Exec(A_ComSpec " /Q /K") ; /Q Turns echo off ; /K keeps cmd open
exec.StdIn.WriteLine("powershell.exe`n" commands "`nExit")
exec.StdIn.Close()
return exec.StdOut.ReadAll()
}

;------------------------------------------------------------------------------
;  = Locate TranscodedWallpaper
; Retrieve the path of the desktop's current wallpaper - source: AutoHotkey.chm::/docs/lib/ComObject.htm#ExWallpaper

Locate_TranscodedWallpaper() {
cchWallpaper := 260

AD := ComObject("{75048700-EF1F-11D0-9888-006097DEACF9}", "{F490EB00-1240-11D1-9888-006097DEACF9}")
wszWallpaper := Buffer(cchWallpaper * 2)
ComCall(4, AD, "ptr", wszWallpaper, "uint", cchWallpaper, "uint", 0x00000002)
Return StrGet(wszWallpaper, "UTF-16")
}

; Output - Wallpaper: C:\Users\USERNAME\AppData\Roaming\Microsoft\Windows\Themes\TranscodedWallpaper

;------------------------------------------------------------------------------
; ChangeLog

/*
v6.01 - 2024.11.20
 * add changelog
 * rearrange/rename/update headings in TOC

v6.02 - 2024.11.20
 * rearrange/rename/update headings in TOC
 * improve comments and other small changes
*/
