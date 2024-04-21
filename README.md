
In `VALUPACK/MUSICTRK` on the English Office 97 CD-ROM, there lives a little-known music generator for PowerPoint. It seems to be based on an early iteration of DirectMusic - likely from the Interactive Music Architecture (IMA) days.

The music files installed to `C:\Interactive Music` are **styles (`.STY`)** and **personalities (`.PER`)**. The STYs are just like in DirectMusic, and DirectMusic Producer can open them just fine. But what DirectMusic later called "chordmaps" are instead "personalities". These "personailities" do include chordmap-like information... but no version of DirectMusic Producer I've found, going back to the DirectX 7 beta, can open these "personailities". The internal file structure of these personalities dramatically differs from DirectMusic chordmaps.

Just like in DirectMusic, the style defines the motifs and the instruments (bands) in a chord-independent way; the personality actually defines probability paths through chord progressions.

This project has several goals:
 - Identify the music generation engine, since it likely predates DirectMusic.
 - Find the application that can natively create and edit personailities.
 - Write a script to convert personalities to DirectMusic chordmaps, natively editable in DirectMusic Producer.

It seems to be possible to use this Add-In [on modern PowerPoint](https://cmpct.info/~calvin/Articles/PowerPointSoundtracks/), which is pretty cool.

# Who else uses this pre-DirectMusic engine?
I expect research conducted on this PowerPoint music generator to transfer to these:
 - *Blood II: The Chosen* (1997)
 - *Shogo: Mobile Armor Division* (1998)
 - Microsoft Music Producer (1997)

# Random Details

## PPMUSIC.PPA
A VBA PowerPoint Add-In that implements the user interface but calls into the DLLs to actually generate the music. The VBA code from this Add-In is extracted and included in this repository, but if you want to edit the Add-In yourself in PowerPoint 97 you can follow these steps:

1. On the machine where you want to edit the Add-In, make sure all PowerPoint 97 instances are closed.
2. Change the following Registry value to True: `HKEY_CURRENT_USER\Software\Microsoft\Office\8.0\PowerPoint\Options\DebugAddins`. (You might also have to search the Registry and change it in other places too.)
3. Re-open PowerPoint and go to the VBA editor. Now the Add-In is visible in the VBA editor!
4. To view the code, enter the password `vbadev`. (Password revealed thanks to [oledump https://blog.didierstevens.com/2020/07/20/cracking-vba-project-passwords/].)

`olevba` doesn't seem to be able to extract the code from this Add-In; I think I just did a manual export from the VBA editor.

## PPMUSAU.DLL
Coming soon!

## PPMUSSET.DLL
Coming soon!
