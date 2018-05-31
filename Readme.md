
# Adhesive Windows

### A  Window Manager for VB.NET Windows Forms Applications.

![Testing Adhesive](http://darn.nl/adhesive/adhesive_test_app2.png)

## Features
 - **Automaticly** re-opens all last used tool windows on startup.
 - Offers named **Layout Presets**, to quickly switch between different jobs. 
 - Makes the windows stick to other application windows (**WinAmp** style).
 - Moves the Tool windows with the Main window, even the **closed** tools.
 - Achives correct "sticky" behaviour on all Windows versions by using '**VisualBounds**' internally.
 - Handles **Aero Auto Snap** events (Window Arrangement, Multitasking) in a conceivable way. 
 - **Easy to Implement** in any VB.NET Windows Forms application.
 - Comes with a fully featured **Test Application** to quickly get the idea, and as a base for further development.

## Implementation

 1. Add the '**AdhesiveWindows**' subdirectory (with **Adhesive.vb** and **NativeMethods.vb**) to your project.
 2. Create one **User Settings** variable; a String called '**Adhesive**'.
 3. Add the reference "**Imports AdhesiveWindows**" and the following code to your Forms:
>     '----- start Adhesive code -----
>     Friend adhesive As Adhesive = Nothing
>     Protected Overrides Sub WndProc(ByRef m As Message)
>         If (adhesive Is Nothing) Then adhesive = New Adhesive(Me)
>         If (adhesive.FilterMessage(m)) Then MyBase.WndProc(m)
>     End Sub
>     Public Overloads Sub Show()
>         Adhesive.OpenToolForm(Me)
>     End Sub
>     '----- end Adhesive code -----

 4. **Done!** Run your solution and be amazed.

## Testing Adhesive
Would you like to check out the Adhesive Windows features? **Great!** Do you have [**GIT**](https://github.com/git-for-windows/git/releaseshttps://github.com/git-for-windows/git/releases/latest "GIT for Windows") installed? Or [**Github Desktop**](https://desktop.github.com/ "Github Desktop")? Then you can clone the Adhesive Test app from [**here**](https://github.com/everheul/AdhesiveWindows.git "Adhesive Test") to a local directory. Otherwise, download the zip and unpack it. Then open the project in Visual Studio, &#x2BC8;**Start** it... and be amazed.
Please remember, reactions from testers, especially on *rare* Windows versions, are crucial for a Windows Manager. So DO let us know about your experiences!

![Testing Adhesive](http://darn.nl/adhesive/adhesive_test_app1.png)

 1. **Start**. The first thing you may notice, when you start the Adhesive Windows Test App, is that it already opens a set of tool windows for you. That's in the default user settings - if you didn't scale your display to 125% or more. If you now move, size, close and/or open some tool windows and then exit, next time you will find everything exactly the way you left it.
 2. **Test Windows.** You can create as many Test Windows as you like, open them, close them, or click on "Reset" to remove them all. The border type of a test window can be changed without changing its "Visual Bounds". A pleasantry. Or?
 3. **Snap or Move.** If one wants to move the main window **without** dragging the tools with, snapping to tool windows instead.. that's completely possible.  Think of a friendly way to implement this feature, it can be very handy.
 4. **Presets.** The default user settings also contain a number of presets, displayed in the "Presets Tool". If you double-click one of the presets, or click the "Apply" button, the layout will change into what it was when the preset was created. You can make your own layout presets and save, overwrite, delete or apply them. Notice that the changes are quite smooth; Adhesive tries not to work against the system.
 5. **Force.** Adhesive calculates the snapping borders beforehand. And while moving or resizing a window, if you get **too close** to a border, the window gets attached to it: it 'snaps' into place. How close exactly can be altered, see what works best for you. **Note:** The snapping borders are straight lines, vertical and horizontal, so there doesn't have to be another window around to snap. Also, two borders can't be too close to each other, that feels sloppy. If that happens, the border from the farthest window is deleted. 
 6. **Gap Size.** Letting the user change the gaps for both X and Y might make up for any 'Form Bound' surprises Microsoft may come up with in the future. But it sure can help you choose a static gap size. Changing the gap size will move the snapped tool windows, to give you an idea, but it won't resize them, so the result may not satisfy completely. Give it a try anyway.
 7. **Aero Movements.** Since Windows Vista, the options to move and size a application window around have been renamed and changed a few times. In Windows 7 one could move a window to the left, top or right, to size it to half- or full-screen. It was called "Automatic Window Arrangement". Windows 10 calls it "Multitasking" and also offers the corners, to have your window quarter-screen sized. Both let you double-click the top-size-border to make it full-height... did you know that one? Admittedly, from the Window Managers point of view, Aero is a real nightmare. Adhesive tries to cope by hiding the tool windows after Aero resizes and by disabling the maxi- or minimize buttons when they potentially could harm the layout. See what you think of the result.

## WinAmp Sticky
If your solution only needs WinAmps "Sticky" effect, you should also check out StickyWindows. Corneliu tried a different approach, apart from the WndProc override, but it works fine:

### StickyWindows ###
Copyright (c)2004 Corneliu I. Tusnea
https://www.codeproject.com/Articles/6045/Sticky-Windows-How-to-make-your-top-level-forms-to

Converted to VB.NET (c)2010 Jason James Newland
http://www.a1vbcode.com/app-5137.asp

Converted to C++ (c) 2017 Thomas Freudenberg
https://github.com/thoemmi/StickyWindows
