# AHK v2 scripts  

[AutoHotkey](https://github.com/Lexikos/AutoHotkey_L/) scripts with examples of commonly used commands, shortcut keys and several small functions for various applications, written in AHK v2. 

I'm not a programmer, but I am learning to script in AHK v2, slowly :)  
Any help would be appreciated. Any suggestions for improving the script code are welcome.  

Standalone scripts for each function will be created soon.

-----------------

### Script No. 1 Showcase - [Link](https://github.com/xypha/AHK-v2-scripts/blob/main/%231%20Showcase.ahk)

CONTENTS
 
  * Set default state of lock keys
  * Toggle protected operating system (OS) files
  * Customise Tray Icon
  * Check & Reload AHK
  * Remap Keys
  * Customise CapsLock
  * Horizontal Scrolling
  * Move Mouse Pointer by pixel
  * Close or Kill an app window
  * Adjust Window Transparency
  * Recycle Bin shortcut
  * Display Off shortcut
  * Add Control Panel Tools to a Menu
  * Change Text Case
  * Wrap Text In Quotes or Symbols
  * Exchange adjacent letters
  * Toggle Window On Top
  * Use cases for Hotstrings
  * Use cases for #HotIf
    * Firefox
    * Telegram
    * Windows Explorer
  * Capitalise the first letter of a sentence
  * Notification Function

<sub>Click on ⏵ to show/hide brief descriptions of above contents.</sub>
<details>
<summary>DESCRIPTIONS</summary>

#### **Set default state of lock keys**
  Set the state of `CapsLock`, `NumLock` and `ScrollLock` to On or Off upon script start.  
  Add this script to system [startup](https://www.howtogeek.com/208224/how-to-add-a-program-to-startup-in-windows/#step-two-create-a-shortcut-in-the-quot-startup-quot-folder-to-add-a-program-to-startup) and set lock-state automatically after you login.

-----------------

#### **Toggle protected operating system (OS) files**
  Show/hide protected operating system files in Windows File Explorer from the script tray menu, as an alternative to navigating the labyrinth of Explorer's Folder Options. Also, a handy check mark is displayed when OS files are shown.

  ![Toggle OS files](https://github.com/xypha/AHK-v2-scripts/assets/12472214/5d409108-ab10-4877-8be5-4c158da140b8)

-----------------
#### **Customise Tray Icon**
  Change the script tray icon (this also changes the default icon in a script's error windows, msg boxes and GUI title icon).
  An example of an icon file to customise the tray icon is included in the folder titled 'icons'.
  
-----------------
#### **Check & Reload AHK**
  A keyboard shortcut to check the script's recent actions using `ListLines`, and another shortcut to `Reload` the script after making changes to it.

-----------------
#### **Remap Keys**
  This section of the script includes methods to disable hardware keys that you don't use or trigger accidentally too often.  
  See examples of methods that allow you to customise a key's function to your needs -  
  * Minimise a window instantly with `ALT + M`, instead of moving your mouse cursor all the way to the "Minimise" button in the title bar.
  * Disable `Insert` key from accidentally triggering when you are trying to press adjacent keys like `Delete` or `Pause/Break`
  * If you don't like **Task view**, remap it's shortcut `Win + Tab` to the invoke the (arguably faster) traditional `AltTab` menu.
  * Working on a laptop and missing the `Page Up/Down`, `Home` and `End` buttons? Remap `RCtrl + Up/Down/Left/Right` button combos to regain the function of missing keys.

-----------------
#### **Customise CapsLock**
  Do you accidentally trigger `CapsLock` when trying to press `a` or `Ctrl + A` or `Alt + Tab`? AutoHotkey can ensure this never happens.  
  Do you want to enable CapsLock briefly and have it automatically turn off after 10 seconds? Check the script for examples.

  * Disable `CapsLock` key from accidentally turning ON _only_.
  * Hit `Shift + CapsLock` keys to turn ON CapsLock-state for customisable duration (10s default, range from 250 milliseconds to 49 days), add a quiet notification in the corner and turn it OFF automatically once the timer expires.
  * Do you want to turn OFF CapsLock sooner? Hit `Esc` key or `CapsLock` (even if this key is disabled).

-----------------
#### **Horizontal Scrolling**
  Whether you have a mouse with or without a [tilt wheel](https://en.wikipedia.org/wiki/Scroll_wheel#Functionality), some applications refuse to scroll horizontally.
  This section of the script demonstrates four methods that simulate tilt wheel actions and enables you to force even the most recalcitrant windows to scroll sideways.

-----------------
#### **Move Mouse Pointer by pixel**
  Use `Win + Numpad` keys to move your mouse pointer with precision, pixel by pixel.  
  Cutomise the shortcut commands to your needs by altering one or more of the following -
  * Starting coordinates (absolute or relative position on Screen, app window or client)
  * Speed of mouse cursor (range of 0 (fastest, instant), 2 (default) or 100 (slowest))
  * Degree of mouse movement (relation position by specifying number of pixels).

-----------------
#### **Close or Kill an app window with shortcuts**
  Close any app or window instantly with a keyboard shortcut `Alt + Right Click` without having to navigate to the 'Close' button in the title bar.  
  This simulates the default `Alt + F4` behaviour (in most apps) or `CTRL + W / Q` action (available in some apps).  

  
  Annoyed by unresponsive windows?  
  Instead of performing a long search in Task Manager, kill the unresponsive window immediately with a shortcut `CTRL + ALT + F4`.

-----------------
#### **Adjust Window Transparency**
  Here is a handy way to work with multiple windows on your PC.  
  * Use mouse keys `CTRL + Shift + Wheel Up/Down` to adjust the transparency of an app or window in customisable increments (Mininum `1` to Maximum `255`).
  * Quickly alter transparency to pre-defined levels in two key presses - Hit `F8` and select an option `1 to 5` in the pop-up menu.

  ![Adjust Window Transparency](https://github.com/xypha/AHK-v2-scripts/assets/12472214/317d7536-fa83-456f-93ee-cfdd3ce1fd8b)

-----------------
#### **Recycle Bin shortcut**
  Do you find yourself opening the Recycle Bin multiple times?  
  Do you want to avoid minimising all windows to go to the desktop or scrolling the navigation pane to find the Recycle Bin icon?  
  
  Here is a single shortcut `CTRL + Delete` that will allow you to do the following -  
  * Open the Recycle Bin when Explorer is not open, or  
  * Navigate to the Recycle Bin when Explorer is open, or  
  * Bring the Recycle Bin Explorer window in the background to the foreground, or  
  * Empty the bin when the Recycle Bin Explorer window is in the foreground.  

  Alternatively, one or more of these actions can be assigned to work with different keyboard shortcuts. Autothotkey is awesome like that :)

-----------------
#### **Display Off shortcut**
  A simple keyboard shortcut `CTRL + Esc` to turn off your monitor.

-----------------
#### **Add Control Panel Tools to a Menu**
  Add items missing in the `Win + X` menu to a customisable pop-up menu triggered by `Win + Shift + X`.

  ![Control Panel Tools Menu](https://github.com/xypha/AHK-v2-scripts/assets/12472214/efe11010-ed29-4605-bd14-8063bb268062)

-----------------
#### **Change Text Case**
  Change any length of text to `lower, UPPER, Sentence, Title or iNVERT` case, in-line through a pop-up menu using a single keyboard shortcut.  
  This section of the script works with special characters such as `é → É` and `Â → â` and is Unicode compatible. Search for `TestString` in the script for a more comprehensive example.

  ![Change the case of text](https://github.com/xypha/AHK-v2-scripts/assets/12472214/e6f3c4dd-0b84-4e71-b2ff-e577fb71d9a8)

-----------------
#### **Wrap Text In Quotes or Symbols**
  Enclose words and numbers in different types of quotation marks or symbols `'',"",(),[],{},``,%%,‘’,“”` in-line using a pop-up menu & shortcut keys.

  ![Wrap Text In Quotes or Symbols](https://github.com/xypha/AHK-v2-scripts/assets/12472214/ed53956b-8a5b-47ed-8b08-16fc72e590fa)

-----------------
#### **Exchange adjacent letters**
  This section of the script allows you to switch the positions of two characters by placing the text cursor (`|`) between them and pressing a keyboard shortcut `ALT + L`. The letters reverse positions - `ab|c` becomes `ac|b`.

-----------------
#### **Toggle Window On Top**
  Keyboard shortcut `ALT + T` enables you to keep a specified window on top of all other windows (except other always-on-top windows) and toggle it back to normal.

-----------------
#### **Use cases for Hotstrings**
 Here are some examples of the wide breath of uses for AutoHotkey hotstrings feature
 * Find & replace text in Clipboard with and without regular expressions(RegEx).
 * Trim clipboard text - remove tabs (`\t`), newline markers (`\r \n`), double spaces (`\s+` or "`  `") with and without RegEx.
 * Type current date and/or time in your referred regional format, or any customisable format really,

-----------------
#### **Use cases for #HotIf**
  Create keyboard shortcuts and text replacement commands tailored to specific windows or apps using the `#HotIf` command. This section of the script includes useful examples of shortcuts for [Firefox](https://www.mozilla.org/en-US/firefox/new/), [Telegram](https://desktop.telegram.org/) and Windows Explorer.
  * Firefox
    * Hit `CTRL + Shift + O` to open library / bookmark manager.
    * Disable `CTRL + Shift + Q` shortcut, to prevent accidentally exiting the browser.
  * Telegram
    * Hit `CTRL + Q` to minimise Telegram Desktop to the tray, instead of quitting.
  * Windows Explorer
    * Change function of `F1` (opens [Help](https://go.microsoft.com/fwlink/?LinkId=2171119) in Edge Browser) to `F2` (rename command).
    * Hit `CTRL + Shift + A` to unselect file(s)/folder(s). Conveniently map the opposite of the default behaviour of `CTRL + A` (select all) to a similar shortcut.
    * Use symbols In file names - Explorer prevents users from adding certain symbols `\/:*?"<>|` in a file name. This section of the script allows you to insert similar Unicode symbols as replacements, organically and automatically as you type.  
    ![Symbols In File Names](https://github.com/xypha/AHK-v2-scripts/assets/12472214/c500bf4c-e16d-4c76-b2d4-384a5d54b83c)
    * Copy full path of folder/file(s) to clipboard using shortcut `CTRL + Shift + C`
    * Copy file name(s) without path to clipboard using shortcut `ALT + N`
    * Copy file name(s) without extension and path to clipboard using shortcut `CTRL + ALT + N`

-----------------
#### **Capitalise the first letter of a sentence**
  Do you find it annoying to correct the capitalisation of the first letter of every sentence while writing? This script does it seamlessly and you might not even notice. The script auto triggers in-line after every `Enter (⏎) NumpadEnter (⏎) dot (.) exclamation (!) and question (?) mark`.

-----------------
#### **Notification Function**
  See an example for a custom alert that replaces the depreciated `Progress` command from AHK v1, using `GUI()` in AHK v2. The example allows for personalised notification `text, duration and position`. Alternatively, a similar use of `ToolTip` is also included.
</details>

-----------------
### Script No. 2 MultiClip - [Link](https://github.com/xypha/AHK-v2-scripts/blob/main/%232%20MultiClip%20v3.ahk)

This script was created as a rudimentary alternative to PhraseExpress's [clipboard management](https://www.phraseexpress.com/doc/features/clipboard-manager/) function with inspiration from v1 scripts - [ClipStep](https://autohotkey.com/board/topic/4567-clipstep-step-through-multiple-clipboards-using-ctrl-x-c-v/) and GeekDrop's [Convert Case Cycle](https://geekdrop.com/content/super-handy-autohotkey-ahk-script-to-change-the-case-of-text-in-line-or-wrap-text-in-quotes).  

The script does the following -  
 * `ClipArr` contains 20 slots by default, but can be customised to any number - from 1 to infinity.
 * Hitting `^c` (Copy / `CTRL + C`) or `^x` (Cut / `CTRL + X`) triggers `OnClipboardChange` and saves text[^1] to a slot in clipboard array (`ClipArr`).
 * Hitting `^c` or `^x` a second time moves previously saved text to the next slot and copies the newly selected text to the first slot in `ClipArr`.
 * This behaviour continues until all 20 slots are filled. When new text is added to `ClipArr` on the 21st occasion or later, the oldest copied text (now in 20th slot) is lost.
 * To view text saved in any of the slots, type `p++` to bring up a pop-up menu. Use any one of the shortcuts[^2] to paste text from corresponding slot.

[^1]: Text includes anything from the clipboard that can be expressed as text, such as filenames (with full path) when files/folders are copied from a Windows File Explorer window via `CTRL + C` or `CTRL + X`. See Script #1 Showcase for examples that show you how to copy full path with or without file extension and path.  
[^2]: Shortcuts correspond to the number/alpabet/symbol prior to `=`. Shortcuts are usually underlined, and consist of numbers from NumPad or number row, and keys from the bottom row of QUERTY keyboard. Customise the shortcuts by altering the character _after_ `&` in lines containing `ClipMenu.Add`. 

![MultiClip](https://github.com/xypha/AHK-v2-scripts/assets/12472214/32329607-bf4e-436b-b115-ce1919ab6bc1)

An earlier version of MultiClip (AHK v1 script) can be found [here](https://www.autohotkey.com/boards/viewtopic.php?p=332658#p332658).
