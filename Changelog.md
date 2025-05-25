# Repo

### 2025.05.25

Files
 * add files to `\Lib\` folder - `TextConverter.ahk`, `KeyHistoryWindow.ahk`, `MouseHistoryWindow.ahk`
 * moved menu icons from `\icons\ClipMenu\` to `\icons\Menu\` because icons will be used for all `Menu()` calls, and not just for MultiClip (formerly, ClipMenu)
 * add tray icons for 4, 5 and 6 to `\icons\Tray\`

### 2024.11.29

Files
 * moved tray icons from `\icons\` to `\icons\Tray\`
 * added `ClipMenu` icons to `\icons\ClipMenu\` and their original/modified files to `\icons\Calendar by Kalash\`
 * renamed `\No-2 MultiClip.ahk` to `\standalone\MultiClip.ahk`. It is feature complete, incorporated into the `\Showcase.ahk` script and hence moved to `\standalone\` folder.
 * renamed `\No-1 Showcase.ahk` to `\Showcase.ahk` since it is now the only main script.
 * created `\Changelog.md` to store all changelogs, instead of individual changelogs within script files

Credits.md
 * updated with information about `ClipMenu` icons and improved formatting

# \Showcase.ahk

### v3.00 - 2025.05.25
 + add variable `loadingErrors` to store error msgs occurring during auto-execute when script is started/reloaded and variable `errorDetails` when needed; and commands that shows these errors when loading is finished
 + add examples for hotkeys and hotstrings to launch programs, and add corresponding useful scripts to `\Lib` folder
 + add `Base64 Encode/Decode` hotstrings and corresponding functions
 + add `Accented Vowels` hotstrings and corresponding function `TextCorrection_menuFn`
 + add global variables `MenuShortcuts` and `MenuIcons_path` to improve functions calling on `Menu()`
 + add global variable `Send_LenLimit` to use SendText mode instead of paste in functions such as `pasteThis`, `ChangeCase_action` and `WrapText_action`
 + add more `#HotIf Apps` like `AHK ToolTip`, `Windows Explorer` > `Desktop` and `Run window` and hotkeys/hotstrings in existing sections such as `Notepad++` and `Edit .ahk`
 + add functions `CatchError_details` and `CatchError_show` to record and display error details when using `Try...Catch` statements
 + add function `Folder_move` to help with `DirMove` commands
 + add section `String Fn` with functions that streamline manipulation of strings in other functions
 + add function `ShowErrExit` to show error msgs in msgbox and exit function
 + add section `Compare Arrays & Remove Duplicates` with functions that improve other parts of script such as `^!+F4` hotkey (which was completely rewritten)
 + add classic Notepad to ahk_group `CapitaliseFirstLetter_optIn`
 + add mpc-hc64.exe to ahk_group `CloseWithQW` and `WrapText_disabled`
 + add networx.exe to ahk_group `CloseWithQW`
 + add function `MenuIcons_add` to simplify adding menu icons
 * improve reloading by adding error msg if reload fails using sleep and msgbox
 * reorganise `Windows Explorer` hotkeys/hotstrings to include Desktop as active window for them (`ahk_class Progman`); and improve several of them
 * reduce `SetTimer () =>` timer from -100ms to -1ms
 * change case-sensitive `==` comparisons to case-insensitive `=` comparison wherever case-sensitivity is not needed
 * renamed ahk_group `CloseWithQW` to `CloseWithEscQW`
 * add extra character `+` to some hotstrings in `ahk_group FileNameSymbols` to prevent false positive match while entering RegEx
 * change default location from `A_MyDocuments` to `A_ScriptDir` in a few places
 * move functions from `Launch explorer or reuse to open path` section to `Windows Explorer Fn` section
 * renamed `ClipArr`, `ClipMenu` and associated `MultiClip` functions, variables and headings to refer to `MultiClip` functionality with a prefix; and imrpoved them
 * renamed functions `MyNotificationGui` to `MyNotification_show` and `EndMyNotif` to `MyNotification_close`; and imrpoved them
 * renamed functions `ToggleOS` to `OSfiles_toggle` and `ToggleOSCheck` to `OSfiles_check`; remove redundant function `CheckRegWrite`; incorporate `WindowsRefreshOrRun` function into `RefreshExplorer` itself to remove redundancy
 * improve `GetKillTitles` function and rename it to `KillAll_titles`; remove redundant function `GetKillTitlesFileList`
 * improve `!t` hotkey by moving relevant commands to separate functions `AlwaysOnTop_enable` and `AlwaysOnTop_disable`
 * improve function `SizeFn` to enable `TB` calculation
 * rename function `FileCreate_Or_Append` to `FileCreate_Overwrite`
 * rename section `Calculator view (classic)` to `#HotIf functions` and move inline functions (such as `MarkletFn` and `CloseIncrSearch`) to this section
 * renamed functions in `Adjust Window Transparency` section with prefix `WinTrans_`; and improve them
 * renamed functions in `Change Text Case` section with prefix `ChangeCase_`
 * renamed functions in `Clipboard Fn` section with prefix `Clip_`
 * renamed functions in `Wrap Text In Quotes or Symbols` section with prefix `WrapText_`; and improve them
 * renamed functions in `URL Encode/Decode` section with prefix `URL_`; and improve them
 * renamed functions in `Print Screen Fn` section with prefix `ScreenSnip_`
 * renamed functions in `Windows Explorer Fn` section with prefix `Explorer_` and `ExplorerTab_` (if function is tab-aware); and improve them
 * renamed functions in `Windows Registry` section with prefix `Registry_`
 * renamed functions in `Control Panel Tools` section with prefix `ControlPanel_`
 * replace excessively complicated functions `Paste_via_clipboard` and `RestoreClip` with `pasteThis_clip`
 * replace `and` with `&&`; `or` with `||` in `If` statements
 * improve function `ToolTipFn` by removing excessively complicated commands
 - remove redundant functions `FocusFileList` and `CaptureFolderPath`
 * use continuation sections to breakdown long lines for readability and maintainability
 * rearrange/rename/update headings in TOC
 * improve comments and other small changes

### v2.18 - 2024.12.11
 * fix `MouseMove` hotkey `#Numpad5` to centre the mouse for any display by creating new function `MouseMove_screenCenter()`
 - remove unnecessary "ahk_id " from `WinTitle` parameter wherever applicable
 * fix `Toggle Window On Top` hotkey `!t` from throwing error when triggered for windows without title/title bar
 * improve function `GetExplorerSize` by removing display tooltip commands and adding `return` command to send results to calling hotkey/function; add new variable `files` to store folder contents
 * improve function `DeleteEmptyFolder` by adding `DirExist` verification; adding `return` command to send results to calling hotkey/function; using the improved `GetExplorerSize` function to determine if folder is empty before recycling
 * fix `Show folder/file size in ToolTip` hotkey `^+s` as consequence of changes to `GetExplorerSize` function such as adding display tooltip commands
 * fix `Delete Empty Folder` hotkey `^+d` as consequence of changes to `DeleteEmptyFolder` function; add `MsgBox` to show final results of operation
 * improve `Extract from folder & delete` hotkey `^+e` by moving majority of commands into new function `ExtractExplorerFolder()`; replace `CaptureFolderPath()` with simpler `CallClipboardVar()`; add `Loop Parse` command to allow extraction of multiple folders by `ExtractExplorerFolder()`; replace `Tooltip` with `MsgBox` to allow for detailed inspection/copying of extraction results and errors
 - remove unnecessary `.*` from `Loop Files` command
 * fix `Trim` commands to include `\t\s\r` characters
 * rearrange/rename/update headings in TOC
 * improve comments and other small changes

### v2.17 - 2024.11.29
 * update version in script file
 * fix trigger hotkeys for `CloseIncrSearch` function by replacing `*` wildcard with `+` Shift
 * rearrange/rename/update headings in TOC
 * improve comments

### v2.16 - 2024.11.29  
 ★ incorporate `MultiClip.ahk` into the script. Consequently, replace some `Send` commands with `PasteThis` function; and add `OnClipboardChange` callbacks to functions under `Clipboard Fn` section
 ★ add `Notepad++` section to `#HotIf Apps`. This adds `CloseIncrSearch` function and `Edit .ahk` hotstrings to script
 ★ add hotstring `regJump+` to call new functions `RegJump` and `ValidRegistryPath`
 + add shortcut `Ctrl + M` to run saved bookmarklet via keyword in Firefox. Consequently, add function `MarkletFn`
 + add hotkey triggers for `NumpadDel` and `Del` in `Windows File Explorer` section to change the focus to the next file when the currently focused file is deleted
 * rename function `CheckControlRegEx` to `FocusFileList` and improve it
 * rename variable `I_Icon` to `path_to_TrayIcon` and change path from `\icons\` to `\icons\Tray\`
 * fix missing `WinActive` in `Locate desktop background` section by adding `ahk_class Progman`
 * rename function `ValidPath` to `ValidExplorerPath` for clarity
 * rename function `GetFolderSize` to `GetExplorerSize` for clarity
 * check windows registry to see if dark mode is enabled before running `ahkDarkMenu`. Clarified comment about manually commenting out `#Include` dark mode files if dark mode is NOT enabled.
 * replace some `FileExist()` commands with `Try...Catch` commands because it allows for generation of error msg with info for rectification
 * improve `ToggleOS` function by renaming variable `Status` to `show_protected_files` and replacing `CheckRegWrite` function with `Try...Catch` command
 - remove `Changelog` section and move it to `\Changelog.md` file
 * rearrange/rename/update headings in TOC
 * improve comments and other small changes
 
### v2.15 - 2024.11.20  
 ★ add `Calculator (classic)` section to #HotIf Apps; along with `checkCalcView` function
 ★ add `KeePass` section to #HotIf Apps
 ★ add "Locate desktop background" shortcut Win + W to `Windows File Explorer` section of #HotIf Apps; along with associated `WallpaperPath_v4` and `nxtBackground` functions
 * replace some mediocre code in "Show folder/file size in ToolTip" section and `OpenFolder` function with `GetExplorerPath` function for faster performance
 * improve `AHK Dark Mode` section by adding registry check before running `ahkDarkMenu`
 * rename "ahk_group NoWrapText" to "ahk_group WrapText_disabled" to maintain consistency
 * commented out `MSPaintApp` from "ahk_group WrapText_disabled" and change `WrapTextFn` to enable wrapping text when inserting/editing text element (ClassNN: RICHEDIT50W1) in classic paint
 * replace `MSPaintApp` with new example `ahk_class CalcFrame` for "ahk_group WrapText_disabled"
 * replace `/` (forward slash) with `.` (fullstop) in `datetime+` hotstring, to maintain consistency with similar hotstrings in "Format Date / Time" section
 * replace `HKCU` with `HKEY_CURRENT_USER` and `HKLM` with `HKEY_LOCAL_MACHINE` wherever applicable
 * improve `CallClipWait` function to return clipboard if successful
 * improve "Clipboard functions" by adding default value (2 seconds) to `secs` variable
 * fix `SnipFromMenu` function by replacing non-functional `PostMessage` command with `Send`
 ! enabled "UTF-8" encoding for `FileAppend` command, instead of the previous "CP0" (system default ANSI) encoding. Might cause error messages if you are using the `FileCreate_Or_Append` function outside this script
 * rearrange/rename/update headings in TOC
 * improve comments and other small changes

### v2.14 - 2024.11.11  
 * change the nature of `Capitalise first letter` section from opt-out to opt-in. Renamed group name from `CapitaliseFirstLetter` to `CapitaliseFirstLetter_optIn` and changed the example.
 * change behaviour of capitalisation when dot/numdot is followed immediately by alphabet - stop inserting Space to allow for trigger-free typing of file extensions
 * renamed group `MediaKeys` to `MediaKeysRestored` for clarity and added comment on UIAccess
 * for users running explorerPatcher.exe, to prevent explorer shortcuts from breaking, change case sensitive matching to RegEx matching for ClassNN, and create new function `CheckControlRegEx`
 * improve `FocusExplorerAddressBar()` to show tooltip on failure
 * improve comments and other small changes

### v2.13 - 2024.11.08  
 ★ add new shortcuts for `AHK Main Window` - cycling views through `Tab and `+Tab` (like browser tabs)
 ★ add shortcut for copying text into a single line when using `Dark Mode - Window Spy`
 ★ add shortcut for `Delete Empty Folder` in Windows File Explorer -- `^+d`
 ★ add shortcut for `Extract from folder & delete` -- `^+e` -- this extracts all items(folders & files) from selected folder to parent folder and deletes the file after emptying it
 ★ add `ahkDarkMenu()` to enable dark mode for `Menu()`
 ★ add several functions that allow querying recycle bin for size and folders/files; display results and generate reports -- see shortcut `^+s` or function `RBinDisplay()`
 ★ add new function `FileCreate_Or_Append` to open existing file If it exists to given path or open a new file and write contents of provided input
 ★ add new function `MsgBox_Custom` to create custom MsgBox and rename 1/2/3 built-in buttons as required
 * rebuild `GetKillTitles` function - add variable to limit size of MsgBox; add file list button If limit is exceeded and function `GetKillTitlesFileList` for it; improve displayed messages
 * rename function and variables within - `FolderSizeFn` to `SizeFn` ; `FolderSizeB` to `SizeB` and so on
 - remove "Size: " text additions in `SizeFn`. These can be added by calling function after returning results
 * improve `Sort` commands by adding numerical sorting to match file sorting in windows file explorer as applicable
 - remove negative numbers for Sleep and SetTimer commands from user-defined functions by using `Abs()` command and pre-assigned sign
 * improve MsgBox commands by changing options from string to numbers (math)
 * change shortcut to open Firefox Homepage from `^+h` to `!h`, and add explanation in comments
 * change "=" to "==" (again) wherever applicable to enable case sensitivity
 * improve `WindowsRefreshOrRun()` - comment out unnecessary `Sleep` command; rearrange/modify remaining commands
 * replace `!` in If commands with `not` to improve readability
 - remove unnecessary variable `errorTxt` from `ValidPath()`
 * rearrange/rename/update headings in TOC
 * improve comments and other small changes

### v2.12 - 2024.10.31  
 - remove `%A_ScriptDir%` from #include commands. It is already built in
 * rename `EncText` to `WrapTextFn`, `WrapTextFromMenu` to `WrapTextMenuSelectionFn`
 * improve all `WrapText` functions by adding Global variables for easy customisation
 * fix `WrapTextFn` by adding `Loop` command, because incorrect removal of leading/trailing characters under the assumption of mixed wrap characters. Note: keep old RegExReplace command in case some users want this behaviour.
 * fix `CallClipboard` incorrect ToolTip after successful copy by moving it prior to sending `^c`
 * fix `ToolTipFn` - failed to assign 1 and 20 to `WhichToolTip` variable, and hence failure to turn Off some ToolTips
 * rearrange/rename some function headings and update TOC
 * improve comments and small changes

### v2.11 - 2024.10.15  
 * rename `Dark ToolTip` section to `AHK Dark Mode` - to include all lib scripts pertaining to dark mode AHK v2
 * change dark mode ToolTip lib file from `ToolTipOptions.ahk` to `SystemThemeAwareToolTip.ahk`
 ★ add `Dark_MsgBox.ahk` and `Dark_WindowSpy` to lib and rename/modify for easier include and tracking
 * fix `AHKname` version increment
 * improve comments

### v2.10 - 2024.10.11  
 * rename file by replacing `#` with `No-` to avoid GitHub conflict with issue numbering
 ★ add `Dark ToolTip` section to adapt `ToolTipFn` function for windows dark mode
 * improve `ToolTipFn` function by removing unnecessary commands, change variable `ToolTipNo` to `WhichToolTip` (to match AHK docs) and change it from `Global` to `static` variable
 - remove `ToolTipOff` function
 * `Process Priority` shortcut changed from 'Win + Z' to `Win + P` (formerly Project shortcut) because Win + Z is newly designated shortcut for Snap Layouts
 - remove unnecessary variable `transformed` from `ConvertSentence` function
 * fix `EncText` function from removing trailing character unintentionally if leading character is not removed and remove unnecessary `StrReplace` command (due to previous `RegExReplace`)
 * use existing `OpenFolder` function to Run explorer command in `Print Screen` section
 * change `winver` to `taskschd.msc` in `Control Panel Tools`
 * rearrange/rename some function headings and update TOC
 * improve comments and small changes

### v2.09 - 2024.09.16  
 ★ add `RefreshExplorer` function to improve `ToggleOSCheck` -- closes Issue #1 (Yipee! My first issue AND closure!)
 + add `Ctrl + G` UnGroup shortcut to `Windows File Explorer` section
 * change `myduration` argument in `MyNotificationGui` function to use negative numbers because negative Sleep is smaller error than forever cycling SetTimer AND to match ToolTipFn; consequently switch negative multiplier from SetTimer to Sleep
 + add `Renaming ahk_exe qbittorrent.exe` to `FileNameSymbols` group
 + add `NoWrapText` group to replace the lone `WinActive` exclusion for `WrapTextMenuFn` function
 * update SetTimer for more versatility in `End auto-execute` section and `MyNotificationGui` function
 * improve `OpenFolder` by moving/changing some commands to new `FocusExplorerAddressBar` function
 * improve `OpenFolder` by adding a new check to see if existing path is not equal to new path; If equal, no further change is necessary, refocus on file list
 * change default notification from `MyNotificationGui` to `ToolTipFn` wherever appropriate -- personal preference shouldn't affect a public script
 * improve `CallClipboard` and `CallClipboardVar` to return error messages when required; and add `ToolTipFn` notification while waiting for clipboard
 * improve `ToolTipFn` to allow for position and add numbering
 - remove `NewThread` function because it is redundant and confusing; replace with `SetTimer`
 + add RegEx url validation and notifications to `UrlEncode` and `UrlDecode`
 * fix `SnipFromMenu` for new version of screenshot v11.2407.3.0 and later (as of 2024.09.16). Old code is still present under block comment
 * rename `PrintScreenExec`to `ScreenshotFileOp` - to reflect the file operations performed by the function and rename variables; add notification on failure
 * update `SnipMenuFn` as per new order of options in snipping tool, open automatically saved file in mspaint
 * improve `GetFolderSize` by moving/changing some commands to new `ValidPath` function and improve ToolTips
 + add new shortcut for `Display Off` section and update existing `Esc` shortcut to `RControl` specifically
 * update `Display Off` section to work with updated shortcuts
 * improve `CheckRegWrite` to use `==` case sensitive operator since registry values can be non-numeric
 * rearrange/rename some functions and update headings
 * improve comments and small changes

### v2.08 - 2024.03.15  
 ★ add `OpenFolder` function and run `Recycle Bin shortcut` through it
 ★ add `CallClipboardVar` function and improve `Exchange adjacent letters` function by using it, instead of calling clipboard multiple times
 + add disabled `Media Keys Group` and `Media Keys Restored` sections
 + add `ahk_class #32770` dialogue box to `FileNameSymbols` group
 + add `NewThread` function and launch `CapsWait` function through it
 * change position of `CapsLock` notification, to be more centred
 * change `!=` to `!==` wherever applicable to enable case sensitivity
 * change `&WinID` to `&id` in `Adjust Window Transparency keys` section, in order to differentiate from `Global` variable `WinID` used by `SetTransMenuFn`
 + add `^+f` shortcut to close find bar in Firefox
 + add `^+h` shortcut to go home in Firefox
 + add `^i` shortcut for `Invert selection` to Windows File Explorer section
 + add `^!s` shortcut for `Show folder size in ToolTip` to Windows File Explorer section
 - remove unnecessary clipboard call from `!n` and `^!n` `Copy file names` hotkeys
 * change `.r+` hotstring to `.r++` in `Find & Replace dot with space (RegEx)` section
 * improve RegEx needles and optimise in `= Find & Replace in Clipboard` section
 * change `MyNotificationGui` colour scheme to white text on dark background (dark mode)
 * improve `WindowsRefreshOrRun` by adding 2s Sleep and launch using `NewThread`
 * change `< 21` to `<= 20` wherever applicable
 * improve `ToolTipFn` by adding `xAxis?, yAxis?` optional parameters
 * replace `Windows version` with `Task scheduler` in `Control Panel Tools`
 * update headings, spelling
 * improve comments and small changes
 * change changelog order for easier access

### v2.07 - 2024.02.20  
 * improve comments

### v2.06 - 2024.02.05  
 + add defaults to `MyNotificationGui` parameters
 - remove default values from all `MyNotificationGui` function calls
 + add defaults to `ToolTipFn` parameters
 - remove default values from all `ToolTipFn` function calls
 - remove unnecessary quotation marks "" for `MyNotificationGui` and `ToolTipFn` parameters
 - remove unnecessary Space or title from `WinWait` commands
 * replace `RegExReplace` with `StrReplace` where possible to improve performance
 - remove unnecessary `ToggleOSCheck` from `ToggleOS(*)`
 * improve `CheckRegWrite` and `ToggleOSCheck` - call `RegRead` only once per `RegWrite`
 ? improve `Recycle Bin shortcut` - replace `WinWait` with `{F6 2}`
 * fix `CaseConvert` - remove `\r` from `caseText` before assigning to `A_Clipboard`
 * fix `EncText` - remove `\r` from `A_Clipboard` before assigning to `TextString` and `TextStringInitial`; instead of single Space, use `&OutputVar` to modify `TextString` and `Len`
 * improve `EncText` - call `A_Clipboard` only once, rename variable `Len1` to `Len`
 - remove unnecessary variable `Len2` from `EncText`
 * improve `CaseConvert` and `EncText` - select text only if string is ≤ 20 characters (change limit as needed), this is to prevent sending large number of keystrokes when these functions are used for big chunks of text
 - remove unnecessary `func` variable from `UrlEncode`
 * improve comments
 * improve changelog - use "fix" instead of "correct/update", use "+" for new additions and "-" for removals, "★" for new functions/sections instead of "*"

### v2.05 - 2024.02.04  
 ★ add `Print Screen` section
 - remove `Telegram` from `CloseWithQW` group - conflict with default behaviour of `Esc`
 + add `MediaInfo in mpc` to `CloseWithQW` group
 + add alternative `PostMessage` to `WinMinimize`
 + add alternative `WinExist` to `MouseGetPos`
 * fix position of parentheses in `WinClose` with `!RButton`
 * fix alternative `^!F4` hotkey - add missing `MouseGetPos`
 - remove unnecessary variable `Process_Name` from alternative `^!F4` hotkey
 * improve `Adjust Window Transparency keys` - call `MouseGetPos` once instead of twice for each mouse key
 * fix `Adjust Window Transparency keys` - `If Trans` statements in `^+WheelDown` hotkey
 * change `A_ThisHotkey` to `ThisHotkey` when applicable for more reliability
 * improve `Recycle Bin shortcut` - add `WinWait` to prevent dropdown explorer bug
 - remove unnecessary variable `SwappedLetters`
 * fix colon replacement with `U+003A` in `Symbols In File Names keys`
 * fix `CheckRegWrite` - uncomment `Exit` command, to stop further execution on `RegWrite` failure
 - remove unnecessary `ToolTip` command from `GetTrans`
 * improve `SetTransByWheel` - add `WinSetTransparent` 255 before setting "Off", and replace `ToolTip`/`SetTimer` combo with `ToolTipFn` and place it after `WinSetTransparent` command
 * fix `SetTransMenuFn` - get `WinID` before executing `SetTransByMenu`
 * improve `SetTransByMenu` - add `WinSetTransparent` "Off" if 255, and add `ToolTipFn`
 * improve comments and update headings

### v2.04 - 2024.01.31  
 * improve remap keys section to show more variations of key names, symbols and formatting
 * replace `HKEY_CURRENT_USER` with `HKCU` in reg keys
 * rename `ShowSuperHidden_Status` to `Status` for future expansion of function
 * condense `ToggleOSCheck` function
 * improve comments and update headings

### v2.03 - 2024.01.30  
 * rename function names with `Func` in the name to `Fn` because `Func` is a class
 * fix `Toggle Window On Top` - change `WinSetTitle` command to apply to known variable `t` instead of `"A"`
 * other minor changes

### v2.02 - 2024.01.29  
 * improve `ListLines` `WinWait` command by using variables

### v2.01 - 2024.01.28  
 * rename `MyNotificationFunc` to `MyNotificationGui`
 * specify "Save As" WinTitle in `FileNameSymbols` group to avoid conflict with apps that use `ahk_class #32770` for other uses, example - Notepad++ find-replace dialogue box
 * remove unnecessary parentheses in `If` commands
 * add exclusion for Notification Center and System tray overflow window for `WinClose` command
 * change `WinKill` command to use active window and add alternative
 * improve continuation section in `^!+F4` hotkey
 * rename `GetTitles` to `GetKillTitles` for more specificity, and move to user-defined functions
 * improve `GetKillTitles` - add padding, `RegExReplace` for `\t \r \n`, add word wrap
 * improve `^!+F4` hotkey to use `ProcessClose` instead of `Run A_ComSpec`
 * rename `SetTrans` to `SetTransByWheel` for more specificity
 * rename `Tool_TipFunc` to `ToolTipFunc`
 * add link to `Display Off shortcut`
 * update `SoundPlay` and `SoundBeep` commands to AHK v2 in `Capitalise first letter` section
 * replace `||` with `or` in single line `If` commands
 * replace `!` with `not` in `If` commands
 * rename `SetTransFunc` to `SetTransByMenu` for more specificity
 * move position of `SetTransByMenu`
 * rename `WrapTextFunc` to `WrapTextFromMenu` for more specificity
 * remove unnecessary code variable from `UrlDecode`
 * rename `ControlPanelFunc` to `ControlPanelSelect` for more specificity
 * utilise `ch` variable in `UrlEncode`
 * some minor changes
 * improve comments and update headings

### v2.00 - 2024.01.27  
 * add changelog
 * add variable `AHKname` to easily update script name and version in template and standalone scripts
 * improve comments

# \standalone\MultiClip.ahk

### v4.12 - 2024.11.29  
 ! enabled "UTF-8" encoding for `FileRead`, `FileAppend` command, instead of the previous "CP0" (system default ANSI) encoding. Might cause error messages if you have previously used this script
 * replace some `FileExist()` commands with `Try...Catch` commands because it allows for generation of error msg with info for rectification
 * check windows registry to see if dark mode is enabled before running `ahkDarkMenu`. Clarified comment about manually commenting out `#Include` dark mode files if dark mode is NOT enabled.
 * increase `LimitClipArr` to 25
 * changed `ClipArr` default slot assignment to show slot number and shortcut key
 * use variable `LimitClipArr` to `Initialise ClipArr hotstrings`
 * rename variable `I_Icon` to `path_to_TrayIcon` and change path from `\icons\` to `\icons\Tray\`
 * rename hotstring - `test++` to `testclip+` and move to `ClipArr testing` section
 + add hotstring - `restoreclip+` to restore clipArr contents from file for reasons (e.g., after running `testclip+` hotstring)
 * fix for `SaveClipArr` function running incorrect number of iterations in Loop and throwing index error
 * fix for `SaveClipArr` function by removing unnecessary trailing delimiter from `ClipArrFile`
 * improve `ClipMenuFn` - store slot shortcuts in variable `ClipShortcuts`; use Loop to add menu items; include `SetIcon` command to add icons (if they exist) to menu items and add warning if icons don't exist or failed to load. Consequently add `ClipMenu_icon_error` function to show error msg
 * fix `ClipTrim` by adding missing "`r" to StrReplace command
 * move some commands in `PasteThis` to separate function `Paste_via_clipboard`, to allow for direct usage if needed
 - remove `Changelog` section and move it to `\Changelog.md` file
 * rearrange/rename/update headings in TOC
 * improve comments and other small changes

### v4.11 - 2024.11.10  
 ★ add `ahkDarkMenu()` to enable dark mode for `Menu()`
 - remove negative numbers for Sleep and SetTimer commands from user-defined functions by using `Abs()` command and pre-assigned sign
 * replace `.Get` commands with `[Index]` when applicable because they are equivalent and "__Item is not called"
 * rearrange/rename/update headings in TOC
 * improve comments and other small changes

### v4.10 - 2024.10.31  
 * correct positioning of `AHK Dark Mode` section as per TOC
 * remove `%A_ScriptDir%` from #include commands. It is already built in
 * add missing `Else` commands to `ClipChanged` function
 * fix `ToolTipFn` - failed to assign 1 and 20 to `WhichToolTip` variable, and hence failure to turn Off some ToolTips
 * improve comments and small changes

### v4.09 - 2024.10.15  
 * rename `Dark ToolTip` section to `AHK Dark Mode` - to include all lib scripts pertaining to dark mode AHK v2
 * change dark mode ToolTip lib file from `ToolTipOptions.ahk` to `SystemThemeAwareToolTip.ahk`
 ★ add `Dark_MsgBox.ahk` and `Dark_WindowSpy` to lib and rename/modify for easier include and tracking
 * improve comments

### v4.08 - 2024.10.11  
 * rename file by replacing `#` with `No-` to avoid GitHub conflict with issue numbering
 ★ add `Dark ToolTip` section to adapt `ToolTipFn` function for windows dark mode
 * improve `ToolTipFn` function by removing unnecessary commands, change variable `ToolTipNo` to `WhichToolTip` (to match AHK docs) and change it from `Global` to `static` variable
 - remove `ToolTipOff` function
 * change `myduration` argument in `MyNotificationGui` function to use negative numbers because negative Sleep is smaller error than forever cycling SetTimer AND to match `ToolTipFn`; consequently switch negative multiplier from SetTimer to Sleep
 * change `ClipMenuFn` shortcut from `p++` to `c++` - more intuitive
 * change `c++` shortcut to `c1+` - more intuitive, old legacy and also match `PasteCStrings` function
 * improve `InsertInClipArr` by removing unnecessary preliminary comparison, preventing addition of empty strings to clipArr and adding conditional ToolTip
 + move `ClipChanged` ToolTip to new function `ClipArrToolTipFn` to improve handling
 * improve `ClipTrim` to show errors and insert error message in `ClipArr`
 * use text mode with `Send` in `PasteThis` function
 - remove `PasteAndSend` and `SendAndPaste` functions - redundant and `ListLines` confusion
 * rearrange/rename some function headings and update TOC
 * improve comments and small changes

### v4.07 - 2024.03.15  
 * change "!=" to "!==" wherever applicable to enable case sensitivity
 * change `MyNotificationGui` colour scheme to white text on dark background (dark mode)
 * improve `ClipChanged` alert ToolTip by removing unnecessary clipboard call and increasing character limit from 600 to 1000
 * replace < 16 with <= 15 in `PasteThis`
 * improve `ToolTipFn` by adding `xAxis?, yAxis?` optional parameters
 * improve comments, spelling and small changes
 * change changelog order for easier access

### v4.06 - 2024.02.20  
 * improve comments

### v4.05 - 2024.02.05  
 + add defaults to `MyNotificationGui` parameters
 - remove default values from all `MyNotificationGui` function calls
 + add defaults to `ToolTipFn` parameters
 - remove default values from all `ToolTipFn` function calls
 - remove unnecessary quotation marks "" for `MyNotificationGui` and `ToolTipFn` parameters
 * improve `PasteThis` - use `Send` command if `pasteText` is less than 16 characters
 * improve comments
 * improve changelog - use "fix" instead of "correct/update", use "+" for new additions and "-" for removals, "★" for new functions/sections instead of "*"

### v4.04 - 2024.02.01  
 * improve comments

### v4.03 - 2024.01.30  
 * rename function names with `Func` in the name to `Fn` because `Func` is a class

### v4.02 - 2024.01.29  
 * rename `MyNotificationFunc` to `MyNotificationGui`
 * improve `ListLines` `WinWait` command by using variables
 * rename `Tool_TipFunc` to `ToolTipFunc`
 - remove unnecessary variable `hkey` in `PasteV` function
 - remove unnecessary variable result in `PasteAll` function
 * improve RegEx in `PasteAll` function
 * rename `ClipTrimFunc` to `ClipTrim`
 - remove unnecessary parentheses in `If` commands
 + add `PasteAndSend` and `SendAndPaste` functions
 * some minor changes
 * improve comments and update headings

### v4.01 - 2024.01.27  
 - remove version from file name
 + add alternative method to populate slots in `ClipMenu`
 - remove unnecessary variable `ClipTrim` from `ClipTrimFunc`
 + add `FileExist` command to `SaveClipArr` to prevent error on first exit

### v4.00 - 2024.01.27  
 + add variable `AHKname` to easily update script name and version in template and standalone scripts
 + add changelog
 * improve comments
