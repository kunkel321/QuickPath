# QuickPath

It's pretty similar to the AHK v1 QuickSwitch. I had suggested similar functionality to the QAP developer last year, and user Horst referred me to QuickSwitch, posted by NotNull, at the voidtools (Everything) forums. More recently I noticed the AutoHotkey forum post by Eureka. Also, the associated GitHub repo is by user gepruts. Presumably all three are the same person.

It's a great little utility and I've been using it steadily for the last year and a half. I wanted an AHK v2 version of it. I attempted to convert the v1 to v2 but was unsuccessful.

Recently, william_ahk's v2 QuickerAccess, and CyberKlabauter's enthusiastic evolution of it, motivated me to feed QuickerAccess into Claude AI and attempt to remake QuickSwitch for AHK v2. Though AI was used, a great deal of human direction was needed. I ended up using NotNull's file-dialog-detection technique, and the functionality of QuickPath is essentially a subset of the functionality of QuickSwitch. It seems to be working well.

Currently, only Win Explorer and Directory Opus file managers are supported. I don't have Win 11, so I can't test it.

QuickPath has a minimal interface. When a Windows 'Open' or 'Save as' type of dialog is opened, whatever (if any) folders are open will by listed in a popup menu. Click a menu item to quickly use that folder as the location in the the Windows dialog. Active DOpus tabs are marked with a little blue arrow icon.

![Screenshot of QuickPath popup menu](https://github.com/kunkel321/QuickPath/blob/main/QuickPath%20screenshot.png)

Tips:
* There is a 2 second timeout, and QuickPath will stop popping-up after this time. So if you want to see it again, press Alt + q.
* Right-click the SysTray icon to set QuickPath to run at Windows start up. There is a checkbox item at the bottom of the R-Click menu.

AutoHotkey Forum Thread: https://www.autohotkey.com/boards/viewtopic.php?f=83&t=134987
