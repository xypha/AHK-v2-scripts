<!-- 
To Do:
* add Print Screen section
* add 'Close With Esc/Q/W'
  * Telegram
* Hit `CTRL + Q` to minimise Telegram Desktop to the tray, instead of quitting.
* add URL Encode/Decode
* update all section based on changed code

; https://github.com/xypha/AHK-v2-scripts/edit/main/README.md
; Last updated 2024.10.11
-->

# AHK v2 scripts  

  Here are some [AutoHotkey](https://github.com/Lexikos/AutoHotkey_L/) scripts written in AHK v2.  
  The script showcase contains a mix of custom shortcut keys and text replacement intended to perform useful tasks, execute commonly used commands and run several small functions.  

  Suggestions for improving the script code are welcome.  

## Script No. 0 Template

  [Link](https://github.com/xypha/AHK-v2-scripts/blob/main/standalone/Template.ahk) to template script.  
  <sub>Last updated (yyyy.mm.dd) - 2024.03.28</sub>

### **MyNotification**  
  
  Create a personalised alert that allows you to customise it's `text, duration and position`.  
  This alert is coded in AHK v2 and utilises the built-in `Gui()` function to display the alert. It is intended to replace the depreciated `Progress` command from AHK v1.  

## Script No. 1 Showcase

  [Link](https://github.com/xypha/AHK-v2-scripts/blob/main/No-1%20Showcase.ahk) to all-in-one script.  
  <sub>Last updated - 2024.10.11</sub>

### Contents

<!--   | Title                          | Standalone                                                                                                         | Last Updated  |
  | :---                           |    :---:                                                                                                           |     ---:      |
  | Set default state of Lock keys | [Link](https://github.com/xypha/AHK-v2-scripts/blob/main/standalone/Set%20default%20state%20of%20lock%20keys.ahk)  | 2024.03.28    |
  | Show/Hide OS files             | [Link](https://github.com/xypha/AHK-v2-scripts/blob/main/standalone/Show%E2%A7%B8Hide%20OS%20files.ahk)            | 2024.03.28    |

  Standalone scripts for each of the below is in the works.  

### Descriptions -->
  <sup>Click on ⏵ to show/hide some of the descriptions.</sup>

#### **Set default state of Lock keys**
  
  Set the default state of `CapsLock`, `NumLock` and `ScrollLock` to On or Off when the script runs.  
  
  Add this script to your system [startup](https://www.howtogeek.com/208224/how-to-add-a-program-to-startup-in-windows/#step-two-create-a-shortcut-in-the-quot-startup-quot-folder-to-add-a-program-to-startup) and set the lock-state automatically after login.  

-----------------
#### **Show/Hide OS files**
  
  <details>
  <summary><sup>Show/Hide</sup></summary>
  
  Show/hide protected operating system (OS) files in Windows File Explorer from the script's tray menu, as an alternative to navigating and changing Explorer's Folder Options.  

  ![Show/Hide OS files](https://github.com/xypha/AHK-v2-scripts/assets/12472214/5d409108-ab10-4877-8be5-4c158da140b8)

  The script checks the current state of the setting on start using Windows Registry and shows a handy check mark in the tray menu when the option to display protected OS files is enabled.  
  
  </details>

-----------------
#### **Customise Tray Icon**
  
  View code that changes the script's tray icon, which also changes the default icon in a script's error windows, msg boxes and GUI title icon.  
  
  To demonstrate the use of an external icon file for this purpose, a few icon files are included in the [/icons](/icons) folder.  
  
-----------------
#### **Check & Reload AHK**
  
  Two Numpad keyboard shortcuts, one to check the script's recent actions using `ListLines`, and another to `Reload` the script.  

-----------------
#### **Remap Keys**
  
  <details>
  <summary><sup>Show/Hide</sup></summary>
  
  See several examples of how to disable hardware keys that you don't use or trigger accidentally too often, or repurpose the function of such keys to your needs -  
  * Disable the `Insert` key from accidentally triggering when you are trying to press adjacent keys like `Delete` or `Pause/Break`.  
  * Prefer `Alt + Tab` over **Task view**? Remap `Win + Tab` shortcut to always invoke the legacy `Alt + Tab` menu.  
  * Are you using a laptop and miss the `Page Up/Page Down/Home/End` buttons? Remap the `RCtrl + Up/Down/Left/Right` button combos to regain the function of the missing keys.  
  * Minimise a window instantly by pressing `ALT + M`, instead of moving your mouse cursor to the title bar to click on the "Minimise" button.  
  
  </details>

-----------------
#### **Customise CapsLock**
  
  <details>
  <summary><sup>Show/Hide</sup></summary>
  
  Do you accidentally trigger `CapsLock` when trying to press `A` or `Shift` or `Alt + Tab`?  
  Do you want to enable CapsLock briefly and have it automatically turn off after 10 seconds?   

  This script can help you do the following -  
  * Disable the `CapsLock` key to prevent it from accidentally turning ON.
  * Hit the `Shift + CapsLock` keys to turn ON CapsLock-state for a customisable duration (10 seconds by default, but can range from 250 milliseconds to 49 days).  
  * Add a quiet notification in the corner when CapsLock is ON and dismiss it automatically once CapsLock is off.
  * If CapsLock is ON, turn it off instantly by hitting the `Esc` key or the `CapsLock` key (even if CapsLock is remapped to never turn ON). 
  
  </details> 

-----------------
#### **Move Mouse Pointer by pixel**
  
  <details>
  <summary><sup>Show/Hide</sup></summary>
  
  Use `Win + Numpad 1-9` keys to move your mouse pointer with precision, pixel by pixel, using AHK's built-in `MouseMove` command.  
  
  Customise the `MouseMove` command to your needs by altering one or more of the following -
  * Starting coordinates (absolute or relative position on screen, app window or client)  
  * Speed of mouse cursor [range of 0 (fastest, instant), 2 (default) or 100 (slowest)]  
  * Degree of mouse movement (change in absolute/relation position by a precise number of pixels)  
  
  </details>

-----------------
#### **Close or Kill an app window with shortcuts**
  
  <details>
  <summary><sup>Show/Hide</sup></summary>
  
  View three methods to close or kill an app or window -
  
  * Send `WinClose` command with `Alt + Right Click` - closes any app or window instantly without having to navigate to the 'Close' button in the title bar. This simulates the default `Alt + F4` behaviour (in most apps) or the `CTRL + W / Q` action (available in some apps).  

  Annoyed by an unresponsive window? Instead of opening Task Manager, try this first -  
  * Send `WinKill` command with `Ctrl + Alt + F4` - kill the unresponsive window forcefully by terminating its process.  
  
  Didn't work? Kill all instances of an app - 
  * Send `ProcessClose` command with `Ctrl + Alt + Shift + F4` - The script shows a warning, with window titles of visible and hidden windows of the unresponsive process and asks for confirmation before terminating the process.  

  ![Kill All Instances Of An App](https://github.com/xypha/AHK-v2-scripts/assets/12472214/f2ecd9b9-74dc-4e4c-b422-440ad3567d65)
  
  </details>

-----------------
#### **Adjust Window Transparency**
  
  <details>
  <summary><sup>Show/Hide</sup></summary>
  
  Here is a handy way to work with multiple windows on your PC.  
  * Use mouse keys `CTRL + Shift + Wheel Up/Down` to increase/decrease the transparency of an app or window. Transparency values range from `1` (invisible) to `255`(opaque) and mouse keys increase/decrease transparency value by 20 (customisable - see variable `Trans +` and `Trans -`).  
  * Quickly set the transparency of an app or window to pre-defined levels in two key presses - Hit `F8` and select an option `1 to 5` from the pop-up menu.

  ![Adjust Window Transparency](https://github.com/xypha/AHK-v2-scripts/assets/12472214/317d7536-fa83-456f-93ee-cfdd3ce1fd8b)
  
  </details>

-----------------
#### **Recycle Bin shortcut**
  
  <details>
  <summary><sup>Show/Hide</sup></summary>
  
  Do you find yourself opening the Recycle Bin multiple times?  
  Do you want to avoid minimising all windows to go to the desktop or scrolling the navigation pane to find the Recycle Bin icon?  
  
  Here is a single shortcut `CTRL + Delete` that will allow you to do the following -  
  * Open the Recycle Bin when Explorer is not open, or  
  * Navigate to the Recycle Bin when Explorer is open, or  
  * Bring the Recycle Bin Explorer window in the background to the foreground, or  
  * Empty the bin when the Recycle Bin Explorer window is in the foreground.  

  Alternatively, assign one or more of these actions to various keyboard shortcuts. AutoHotkey is awesome like that :)
  
  </details>

-----------------
#### **Display Off shortcut**
  
  Use a keyboard shortcut `CTRL + Esc` to turn off your monitor.

-----------------
#### **Add Control Panel Tools to a Menu**
  
  <details>
  <summary><sup>Show/Hide</sup></summary>
  
  Add items missing in the `Win + X` menu to a customisable pop-up menu triggered by `Win + Shift + X`.

  ![Control Panel Tools Menu](https://github.com/xypha/AHK-v2-scripts/assets/12472214/efe11010-ed29-4605-bd14-8063bb268062)
  
  </details>

-----------------
#### **Change Text Case**
  
  <details>
  <summary><sup>Show/Hide</sup></summary>
  
  Change any length of text to `lower, UPPER, Sentence, Title or iNVERT` case, in-line through a pop-up menu using a single keyboard shortcut.  
  This section of the script works with special characters such as `é → É` and `Â → â` and is Unicode compatible. Search for `TestString` in the script for a more comprehensive example.

  ![Change the case of text](https://github.com/xypha/AHK-v2-scripts/assets/12472214/e6f3c4dd-0b84-4e71-b2ff-e577fb71d9a8)
  
  </details>

-----------------
#### **Wrap Text In Quotes or Symbols**
  
  <details>
  <summary><sup>Show/Hide</sup></summary>
  
  Enclose words and numbers in different types of quotation marks or symbols `'',"",(),[],{},``,%%,‘’,“”` in-line using a pop-up menu & shortcut keys.

  ![Wrap Text In Quotes or Symbols](https://github.com/xypha/AHK-v2-scripts/assets/12472214/ed53956b-8a5b-47ed-8b08-16fc72e590fa)
  
  </details>

-----------------
#### **Exchange adjacent letters**
  
  This section of the script allows you to switch the positions of two characters by placing the text cursor (`|`) between them and pressing a keyboard shortcut `ALT + L`.  
  The letters reverse positions - `ab|c` becomes `ac|b`.  

-----------------
#### **Toggle Window On Top**  
  
  Keyboard shortcut `ALT + T` enables you to keep a specified window on top of all other windows (except other always-on-top windows) and toggle it back to normal.  

-----------------
#### **Process Priority**  
  
  <details>
  <summary><sup>Show/Hide</sup></summary>
  
  Hit `Win + P` to select and change the priority level of a process.  

  ![Set Priority](https://github.com/xypha/AHK-v2-scripts/assets/12472214/2d0fd2cc-8c5a-4c43-9afc-599ac5aebd56)
  
  </details>

-----------------
#### **Use cases for Hotstrings**
  
  <details>
  <summary><sup>Show/Hide</sup></summary>
  
 Here are some examples of the wide breadth of uses for the AutoHotkey hotstrings feature.  
 * Find & replace text in Clipboard with and without regular expressions(RegEx).  
 * Trim clipboard text - remove tabs (`\t`), newline markers (`\r \n`) and double spaces (`\s+` or "`  `") with or without RegEx.  
 * Type the current date and/or time in your preferred format, regional or otherwise.  
  
  </details>

-----------------
#### **Use cases for #HotIf**
  
  <details>
  <summary><sup>Show/Hide</sup></summary>
  
  Create keyboard shortcuts and text replacement commands tailored to specific windows or apps using the `#HotIf` command. This section of the script includes useful examples of shortcuts for [Firefox](https://www.mozilla.org/en-US/firefox/new/), [Telegram](https://desktop.telegram.org/) and Windows Explorer.  
  * Firefox
    * Hit `CTRL + Shift + O` to open library / bookmark manager.  
    * Don't want to accidentally exit the browser with the `CTRL + Shift + Q` shortcut? Disable it.  
  * Windows Explorer
    * Change function of `F1` (opens [Help](https://go.microsoft.com/fwlink/?LinkId=2171119) in Edge Browser) to `F2` (rename command).  
    * Hit `CTRL + Shift + A` to unselect file(s)/folder(s). Conveniently map the opposite of the default behaviour of `CTRL + A` (select all) to a similar shortcut.  
    * Copy the full path of one or more folder/file(s) to the clipboard using the shortcut `CTRL + Shift + C`.  
    * Copy folder/file name(s) without their path to the clipboard using the shortcut `ALT + N`.  
    * Copy file name(s) without their extension and path to the clipboard using the shortcut `CTRL + ALT + N`.  
  
  </details>

-----------------
#### **Capitalise the first letter of a sentence**  
  
  Do you find it annoying to correct the capitalisation of the first letter of every sentence while writing?  
  This script does it so seamlessly that you might not even notice. The script auto triggers in-line after every `Enter (⏎) NumpadEnter (⏎) dot (.) exclamation (!) and question (?) mark`.

-----------------
#### **Horizontal Scrolling**
  
  Whether you have a mouse with or without a [tilt wheel](https://en.wikipedia.org/wiki/Scroll_wheel#Functionality), some applications refuse to scroll horizontally.  
  
  Here are five methods that simulate tilt wheel actions and force even the most recalcitrant windows to scroll sideways.

-----------------
#### **Symbols In File Names**
  
  <details>
  <summary><sup>Show/Hide</sup></summary>
  
  Windows File Explorer prevents users from adding certain symbols `\/:*?"<>|` in a file name.
  
  This script allows you to insert similar Unicode symbols as replacements, organically and automatically as you type.  

  ![Symbols In File Names](https://github.com/xypha/AHK-v2-scripts/assets/12472214/c500bf4c-e16d-4c76-b2d4-384a5d54b83c)
  
  </details>

-----------------
#### **Notification Function**  
  
  Create a personalised alert programmed in AHK v2 that replaces the depreciated `Progress` command from AHK v1.  
  The alert uses the `GUI()` function and allows you to customise the notification's `text, duration and position`.  
  Alternatively, a similar use of `ToolTip` is also included.


## Script No. 2 MultiClip
  
  [Link](https://github.com/xypha/AHK-v2-scripts/edit/main/No-2%20MultiClip.ahk) to script.  
  <sub>Last updated - 2024.10.11</sub>  

  This script was created as a rudimentary alternative to PhraseExpress's [clipboard management](https://www.phraseexpress.com/doc/features/clipboard-manager/) feature with inspiration from several v1 scripts.  

  The script does the following -  
 * Clipboard array (`ClipArr`) contains 20 slots that automatically save text[^1] when Windows Clipboard is used. The number of slots is customisable to any number (see variable `LimitClipArr`).  
 * Hitting `^x` (`CTRL + X` / Cut) or `^c` (`CTRL + C` / Copy) triggers AHK's `OnClipboardChange` and saves text to a slot in `ClipArr`.  
 * Hitting `^x` or `^c` a second time moves previously saved text to the next slot and copies the newly selected text to the first slot.  
 * This behaviour continues until the 20 slots are filled. After which, when new text is added to `ClipArr` on the 21st occasion or later, the oldest copied text (now in the 20th slot) is lost.  
 * To view text saved in any of the slots, type `c++` to bring up a pop-up menu. Use any one of the shortcuts[^2] to retrieve text from the corresponding slot.  
 * Retrieve text directly from a slot using hotstring `v{slot-number}+`. For example, from the test array[^3], typing `v2+` retrieves `b2` and `v20+` retrieves `t20`.  
 * Retrieve text from multiple slots using hotstring `c{slot-numbers}+`. For example, from the test array, typing `c2+` retrieves text from slots 1 & 2 (`a1-b2`) and `c10+` retrieves text from slots 1 to 10 (`a1-b2-c3-d4-e5-f6-g7-h8-i9-j10`).  
 * `ClipArr` slots persist between restarts, unlike Windows Clipboard. Slot contents are saved to a file upon script exit and loaded back from the file when the script starts.   
 * Old script contents can be retrieved by restoring `ClipArrFile.txt` from Recycle Bin.

  MultiClip pop-up menu `c++` with test array -  
  
  ![MultiClip pop-up menu `c++` with test array](https://github.com/xypha/AHK-v2-scripts/assets/12472214/32329607-bf4e-436b-b115-ce1919ab6bc1)
 

  [^1]: Saves **_only_** text - this includes anything from the clipboard that can be expressed as text, such as filenames (with full path) when files/folders are copied from a Windows File Explorer window via `CTRL + C` or `CTRL + X`. See Script No-1 Showcase for examples that show you how to copy the full path with or without file extension and path.  
  [^2]: Shortcuts correspond to the number/alphabet/symbol before `=`. Shortcuts are usually underlined and consist of numbers from Numpad or number row and keys from the bottom row of the QWERTY keyboard. Customise the shortcuts by altering the character _after_ `&` in lines containing `ClipMenu.Add`.  
  [^3]: Test array with 20 slots containing alphanumerical text `ClipArr := ["a1", "b2", "c3", "d4", "e5", "f6", "g7", "h8", "i9", "j10", "k11", "l12", "m13", "n14", "o15", "p16", "q17", "r18", "s19", "t20"]`.  
