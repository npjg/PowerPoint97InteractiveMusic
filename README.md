# PowerPoint Custom Soundtrack

Calvin Buckley has a great summary of the [DirectMusic-based music generator https://cmpct.info/~calvin/Articles/PowerPointSoundtracks/] included with PowerPoint 97 and  still supported on modern PowerPoint.

The VBA Add-In and DLL behind it were mentioned, but only the references in the VBA project were shown.
I was curious what was actually inside the VBA Add-In, so I endeavored to open it up - first in the PowerPoint 97 VBA editor, and second with modern software.

## Installation
First, of course, you must install the Custom Soundtracks Add-In. (It's at `VALUPACK/MUSICTRK` on the English Office 97 CD-ROM.) 
 - C:\Interactive Music\
 - C:\Program Files\Microsoft Office

## In Office 97 VBA Editor

Make sure all PowerPoint 97 instances are closed, and then change the following Registry value to True: `HKEY_CURRENT_USER\Software\Microsoft\Office\8.0\PowerPoint\Options\DebugAddins`. (You might also have to search the Registry and change it in other places too.)

Then when you open PowerPoint and load the Add-In, you can view it in the VBA editor. The VBA project is also password-protected, but [oledump https://blog.didierstevens.com/2020/07/20/cracking-vba-project-passwords/] revealed the password was simply `vbadev`.

## With `olevba`
`olevba` cannot extract the code from `PPMUSIC.PPA`
