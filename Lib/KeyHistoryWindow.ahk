; KeyHistoryWindow.ahk by lexikos
; Source: https://autohotkey.com/boards/viewtopic.php?f=6&t=26059
; Change Log:
; 2022.05.26 - modified by tuzi
;     Fixed incomplete display of history content.
;     Fixed flicker.
;     Added title and tips.
;     Can run in the background.
; Last checked: 2025.05.24

#NoEnv
#Persistent
#InstallKeybdHook
hHookKeybd := DllCall("SetWindowsHookEx", "int", 13 ; WH_KEYBOARD_LL = 13
    , "ptr", RegisterCallback("Keyboard")
    ; hMod is not required on Win 7, but seems to be required on XP even
    ; though this type of hook is never injected into other processes:
    , "ptr", DllCall("GetModuleHandle", "ptr", 0, "ptr")
    , "uint", 0, "ptr") ; dwThreadId
#KeyHistory(15)  ; Although it starts with #, it is actually a function

OnMessage(0x200, "WM_MOUSEMOVE")
OnMessage(0x201, "WM_LBUTTONDOWN")
OnMessage(0x100, "WM_KEYDOWN")

Menu, Tray, NoStandard
Menu, Tray, Add, Show, MenuHandler
Menu, Tray, Add, Freeze, MenuHandler
Menu, Tray, Add, Exit, MenuHandler
Menu, Tray, Default, Show

; WS_EX_COMPOSITED = E0x02000000 & WS_EX_LAYERED = E0x00080000 -> Double Buffer
Gui, +LastFound -DPIScale +AlwaysOnTop +E0x02000000 +E0x00080000 +HwndhGui
Gui, Margin, 10, 10
Gui, Font,, Lucida Console
Gui, Add, Text, vKHT +0x100, % Format("{:-4}{:-5}{:-7}{:-9}{:-19}{:-30}", "VK", "SC", "Flags", "Elapsed", "Key", "Extra")
Gui, Add, Text, vKH, % Format("{:-4}{:-5}{:-7}{:-9}{:-19}{:-30}", "00", "000", "ea!u", "1000.00", "Browser_Favorites", "KEY_IGNORE_ALL_EXCEPT_MODIFIER")
GuiControlGet, KHT, Pos
GuiControlGet, KH, Pos
GuiControl,, KH  ; clear dummy sizing text
gosub Resize
return

#MaxThreadsBuffer, On
!WheelUp::
!WheelDown::
#MaxThreadsBuffer, Off
    history_size := #KeyHistory() + ((A_ThisHotkey="!WheelUp") ? +1 : -1)
    #KeyHistory(history_size>0 ? history_size : 1)
    ; Delay resize to improve hotkey responsiveness.
    SetTimer, Resize, -10
return

Resize:
    ; Resize label to fit key history.
    gui_h := KHH*#KeyHistory()
    GuiControl, Move, KH, h%gui_h%
    gui_h += KHY + 10

    Gui, +LastFound
    ; Determine visibility.
    WinGet, style, Style
    gui_visible := style & 0x10000000

    ;Gui, Show, % "AutoSize NA " (gui_visible ? "" : "Hide")
    ;** Not used because we need to know the previous height,
    ;   and its simpler to resize manually.

    ; Determine current position and height.
    WinGetPos, gui_x, gui_y, , gui_h_old
    ; Use old height to determine if we should reposition, *only when shrinking*.
    ; This way we can move the GUI somewhere else, and the script won't reposition it.
    ;if (gui_h_old < gui_h)
    ;    gui_h_old := gui_h
    ; Determine working area (primary screen size minus taskbar.)
    SysGet, wa_, MonitorWorkArea

    SysGet, twc_h, 51 ; SM_CYSMCAPTION
    SysGet, bdr_h, 8  ; SM_CYFIXEDFRAME
    if (!gui_visible)
    {
        gui_x = 72 ; Initially on the left side.
        gui_y := wa_bottom-(gui_h+twc_h+bdr_h*2+10)
    }
    else
    {   ; Move relative to bottom edge when closer to the bottom.
        if (gui_y+gui_h//2 > (wa_bottom-wa_top)//2)
            gui_y += gui_h_old-(gui_h+twc_h+bdr_h*2)
    }
    Gui, Show, x%gui_x% y%gui_y% h%gui_h% NA, Key History
return

GuiSize:
    if (A_EventInfo = 1)
        Gui, Hide
return

MenuHandler:
    switch A_ThisMenuItem
    {
        case "Freeze":
            if (A_IsPaused)
            {
                WinSetTitle, ahk_id %hGui%, , Key History
                Menu, Tray, UnCheck, Freeze
            }
            else
            {
                WinSetTitle, ahk_id %hGui%, , Key History - Freezed
                Menu, Tray, Check, Freeze
            }
            Pause Toggle

        case "Show":
            Gui, Show, , Key History - Freezed
            if (!A_IsPaused)
            {
                Menu, Tray, Check, Freeze
                Pause On
            }

        case "Exit":
            ExitApp
    }
return

Keyboard(nCode, wParam, lParam) {
    global KeyBuffer
    static sz := 16 + A_PtrSize

    Critical

    if KeyHistory(1, vk, sc, flags)
        && NumGet(lParam + 0, "uint") = vk
        && NumGet(lParam + 4, "uint") = sc
        && NumGet(lParam + 8, "uint") = flags
        buf_max := 0 ; Don't show key-repeat.
    else
        buf_max := #KeyHistory()

    if (buf_max > 0)
    {
        ; Push older key events to the back.
        if (buf_max > 1)
            DllCall("RtlMoveMemory", "ptr", &KeyBuffer + sz, "ptr", &KeyBuffer, "ptr", buf_max * sz)
        ; Copy current key event to the buffer.
        DllCall("RtlMoveMemory", "ptr", &KeyBuffer, "ptr", lParam, "ptr", sz)

        ; "gosub Show" slows down the keyboard hook and causes problems, so use a timer.
        SetTimer, Show, -20
    }

    return DllCall("CallNextHookEx", "ptr", 0, "int", nCode, "ptr", wParam, "ptr", lParam, "ptr")
}

KeyHistory(N, ByRef vk, ByRef sc, ByRef flags:=0, ByRef time:=0, ByRef elapsed:=0, ByRef info:=0)
{
    global KeyBuffer
    static sz := 16 + A_PtrSize

    if N is not integer
        return false
    buf_max := #KeyHistory()
    if (N < 0)
        N += buf_max + 1
    if (N < 1 or N > buf_max)
        return false

    vk    := NumGet(KeyBuffer, (N-1)*sz, "uint")
    sc    := NumGet(KeyBuffer, (N-1)*sz+4, "uint")
    flags := NumGet(KeyBuffer, (N-1)*sz+8, "uint")
    time  := NumGet(KeyBuffer, (N-1)*sz+12, "uint")
    info  := NumGet(KeyBuffer, (N-1)*sz+16)
    elapsed := time - ((time2 := NumGet(KeyBuffer, N*sz+12, "uint")) ? time2 : time)

    switch info
    {
        case 0xFFC3D44F: info := "KEY_IGNORE"
        case 0xFFC3D44E: info := "KEY_PHYS_IGNORE"
        case 0xFFC3D44D: info := "KEY_IGNORE_ALL_EXCEPT_MODIFIER"
    }

    return (vk or sc)
}

#KeyHistory(NewSize="")
{
    global KeyBuffer
    static sz := 16+A_PtrSize
    ; Get current history length.
    if (NewSize="")
        return (cap:=VarSetCapacity(KeyBuffer)//sz)>0 ? cap-1 : 0
    if (NewSize)
    {
        new_cap := (NewSize+1)*sz
        cap := VarSetCapacity(KeyBuffer)
        if (cap > new_cap)
            cap := new_cap
        VarSetCapacity(old_buffer, cap)
        ; Back up previous history.
        DllCall("RtlMoveMemory", "ptr", &old_buffer, "ptr", &KeyBuffer, "ptr", cap)

        ; Set new history length.
        VarSetCapacity(KeyBuffer, 0) ; FORCE SHRINK
        VarSetCapacity(KeyBuffer, new_cap, 0)

        ; Restore previous history.
        DllCall("RtlMoveMemory", "ptr", &KeyBuffer, "ptr", &old_buffer, "ptr", cap)

        ; (Remember N+1 key events to simplify calculation of the Nth key event's elapsed time.)
        ; Put tick count so the initial key event has a meaningful value for "elapsed".
        NumPut(A_TickCount, KeyBuffer, 12, "uint")
    }
    else
    {   ; Clear history entirely.
        VarSetCapacity(KeyBuffer, 0)
    }
}

GetKeyFlagText(flags)
{
    return ((flags & 0x1) ? "e" : "") ; LLKHF_EXTENDED
        . ((flags & 0x10) ? "a" : "") ; LLKHF_INJECTED (artificial)
        . ((flags & 0x20) ? "!" : "") ; LLKHF_ALTDOWN
        . ((flags & 0x80) ? "u" : "") ; LLKHF_UP (key up)
}

; Gets readable key name, usually identical to the name in KeyHistory.
GetKeyNameText(vkCode, scanCode, isExtendedKey)
{
    return GetKeyName(format("vk{1:02x}sc{3}{2:02x}", vkCode, scanCode, isExtendedKey))
    /* ; For older versions of AutoHotkey:
    ; My Right Shift key shows as vk161 sc54 isExtendedKey=true.  For some
    ; reason GetKeyNameText only returns a name for it if isExtendedKey=false.
    if vkCode = 161
        return "Right Shift"

    VarSetCapacity(buffer, 32, 0)
    DllCall("GetKeyNameText"
        , "UInt", (scanCode & 0xFF) << 16 | (isExtendedKey ? 1<<24 : 0) ;| 1<<25
        , "Str", buffer
        , "Int", 32)

    return buffer
    */
}

Show:
    SetFormat, FloatFast, .2
    SetFormat, IntegerFast, H
    text =
    buf_size := #KeyHistory()
    Loop, % buf_size
    {
        if (KeyHistory(buf_size+1-A_Index, vk, sc, flags, time, elapsed, info))
        {
            keytext := GetKeyNameText(vk, sc, flags & 0x1)

            if (elapsed < 0)
                elapsed := "#err#"
            else
                dt := elapsed/1000.0

            ; AHK-style SC
            sc_a := sc
            if (flags & 1)
                sc_a |= 0x100, flags &= ~1
            sc_a := SubStr("000" SubStr(sc_a, 3), -2)
            vk_a := SubStr(vk+0, 3)
            if (StrLen(vk_a)<2)
                vk_a = 0%vk_a%
            StringUpper, vk_a, vk_a
            StringUpper, sc_a, sc_a

            flags := GetKeyFlagText(flags & ~0x1)

            text .= Format("{:-4}{:-5}{:-7}{:-9}{:-19}{:-30}`n", vk_a, sc_a, flags, dt, keytext, info)
        }
    }
    GuiControl,, KH, % text
Return

GuiClose:
ExitApp

WM_KEYDOWN()
{
    if A_Gui
        return true
}

WM_MOUSEMOVE()
{
    if (A_GuiControl="KHT")
        ToolTip In Flags`ne%A_Tab%=%A_Tab%Extended`na%A_Tab%=%A_Tab%Artificial`n!%A_Tab%=%A_Tab%Alt-Down`nu%A_Tab%=%A_Tab%Key-Up
    else
        ToolTip
}

WM_LBUTTONDOWN(wParam, lParam)
{
    global text
    StringReplace, Clipboard, text, `n, `r`n, All
}