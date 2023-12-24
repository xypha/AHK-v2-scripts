# AHK v2 scripts  

[AutoHotkey](https://github.com/Lexikos/AutoHotkey_L/) scripts with examples of commonly used commands and shortcut keys with several small functions, written in AHK v2. 

Any suggestions for improving script code are welcome.

### Script #1 - [Link](https://github.com/xypha/AHK-v2-scripts/blob/main/AHK%20v2%20%231.ahk)  

- **Toggle protected operating system (OS) files** -\
  Show/hide protected operating system files in Windows File Explorer from AHK tray menu, as an alternative to navigating the labyrinth of Explorer's Folder Options. Also, a handy check mark is displayed when OS files are shown.

  ![Toggle OS files](https://github.com/xypha/AHK-v2-scripts/assets/12472214/5d409108-ab10-4877-8be5-4c158da140b8)

- **Tray Icon** -\
  Change script tray icon (which also changes the default icon in a script's error windows, msgboxes and GUI title icon).
  Example of icon file to customise the tray icon is included in folder titled 'icons'.
  
- **Check & Reload AHK** -\
  Create shortcuts to check the functions of a script using `ListLines`, and to `Reload` a script after making changes to it.

- **Remap Keys** -\
  Script includes methods to disable keys that you don't use or trigger accidentally too often; or see examples of ways to change a key's function to suit your needs, such as pressing `ALT + m` keys to minimise a window, instead of moving your mouse cursor to the "Minimize" button in the title bar. 

- **CapsLock** -\
  Do you accidentally trigger `CapsLock` when trying to press `a` or `Ctrl + A` or `Alt + Tab`? AutoHotkey can ensure this never happens.
  Do you want to temporarily enable CapsLock and have it automatically turn off after 10 seconds? Check the script for an example.
  
- **Control Panel Tools Menu** -\
  Add items missing in `Win + X` menu to a customizable AutoHotkey Menu triggered by `Win + Shift + X`.
  
  ![Control Panel Tools Menu](https://github.com/xypha/AHK-v2-scripts/assets/12472214/efe11010-ed29-4605-bd14-8063bb268062)

- **Automatically capitalize the first letter of a sentence** -\
  In-line using InputHook [auto triggers after every `Enter NumEnter . ! ?`].
  
- **Change the case of text** -\
  Pop up menu that allows you to change the text to `lower, UPPER, Sentence, Title or iNVERT` case, in-line using a pop-up menu.
  This works even with special characters such as `é → É` and `Â → â` (unicode compatible). Search for `TestString` in the script for a more comprehensive example.

  ![Change the case of text](https://github.com/xypha/AHK-v2-scripts/assets/12472214/e6f3c4dd-0b84-4e71-b2ff-e577fb71d9a8)
  
- **HasVal Function** -\
  Check if value is already in an array and retrieve it's `Index` if present.
  
- **Notification Function** -\
  See an example for a custom alert that replaces the depreceated `Progress` command from AHK v1, using `GUI()` in AHK v2. The example allows for personalised notification `text, duration and position`. Alternatively, a similar use of `ToolTip` is also included.
  
- **Wrap Text In Quotes** -\
  Enclose words and numbers in different types of quotation marks or symbols `'',"",(),[],{},``,%%,‘’,“”` in-line using a pop-up menu & shortcut keys.

  ![Wrap Text In Quotes](https://github.com/xypha/AHK-v2-scripts/assets/12472214/f21821c3-028a-4327-974a-6cb62726d297)
