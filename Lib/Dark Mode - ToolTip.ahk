/* Source: https://raw.githubusercontent.com/nperovic/SystemThemeAwareToolTip/main/SystemThemeAwareToolTip.ahk

Last checked: 2024.12.23
Last update: 2024.03.20

Modifications by xypha -
    * 2024.12.23 - add ListLines
*/

#requires AutoHotkey v2

class SystemThemeAwareToolTip
{
    static IsDarkMode => !RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "AppsUseLightTheme", 1)

    static __New()
    {

        if this.HasOwnProp("HTT") || !this.IsDarkMode
            return

        GroupAdd("tooltips_class32", "ahk_class tooltips_class32")

        this.HTT        := DllCall("User32.dll\CreateWindowEx", "UInt", 8, "Ptr", StrPtr("tooltips_class32"), "Ptr", 0, "UInt", 3, "Int", 0, "Int", 0, "Int", 0, "Int", 0, "Ptr", A_ScriptHwnd, "Ptr", 0, "Ptr", 0, "Ptr", 0)
        this.SubWndProc := CallbackCreate(TT_WNDPROC,, 4)
        this.OriWndProc := DllCall(A_PtrSize = 8 ? "SetClassLongPtr" : "SetClassLongW", "Ptr", this.HTT, "Int", -24, "Ptr", this.SubWndProc, "UPtr")

        TT_WNDPROC(hWnd, uMsg, wParam, lParam)
        {
            ListLines 0     ; Omit subsequently-executed lines from the history

            static WM_CREATE := 0x0001

            if (this.IsDarkMode && uMsg = WM_CREATE)
            {

                SetDarkToolTip(hWnd)

                if (VerCompare(A_OSVersion, "10.0.22000") > 0)
                    SetRoundedCornor(hWnd, 3)
            }

            ListLines 1     ; Include subsequently-executed lines in the history

            return DllCall(This.OriWndProc, "Ptr", hWnd, "UInt", uMsg, "Ptr", wParam, "Ptr", lParam, "UInt")
        }

        SetDarkToolTip(hWnd) => DllCall("UxTheme\SetWindowTheme", "Ptr", hWnd, "Ptr", StrPtr("DarkMode_Explorer"), "Ptr", StrPtr("ToolTip"))

        SetRoundedCornor(hwnd, level:= 3) => DllCall("Dwmapi\DwmSetWindowAttribute", "Ptr" , hwnd, "UInt", 33, "Ptr*", level, "UInt", 4)
    }

    static __Delete() => (this.HTT && WinKill("ahk_group tooltips_class32"))

}