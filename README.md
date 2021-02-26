# CShauto

CShauto is an RPA C# library that Provide some methods to manipulate Images, Mouse, Keyboard

## Installation

In order to use the library you need to install the following nuGet packages.

```bash
Install-Package CShauto -Version 1.0.1
```

Then you need to add some references to the project, for this just
right click in references and then click on Add reference and add the following ones in Assemblies menu

```bash
 System.Drawing;
 System.Windows.Forms;
```

Finally add this reference in COM menu
```bash
 Microsoft Internet Controls
```

## Usage

For use is library just add the following
```C#
using CShauto;
```

Here is an example of how to open the remote desktop dialog
---
```C#

//Open the remote desktop windows
KeyBoard.Shortcut(Shortcuts.WINDOW_RUN);
Helpers.Wait(2000);
KeyBoard.Write("mstsc");
KeyBoard.Press(KEYCODE.ENTER);
Window remoteDesktop = Window.GetWindowWithPartialName("remote desktop connection", attempts: 10);
remoteDesktop.Focus();

```

As you can see here we use the **Shortcut** method 
to press **window key + R key** and open the run windows. Then add a **Wait** just in case it takes 
a time to open, then type **mstc** witch is the command to open remote desktop dialog.
After that we just create a window variable that store the remote desktop windows that we just open. Notice that the method **GetWindowWithParcialName** is not case sensitive, Also we set a couple of tries to look for the window with **attempts** param, just in case it takes a while to open, after that we just focus the window with the **Focus** method.

Manipulate images:
---
```C#
//Manipulate images
Rect region = ScreenRegions.Left();
Rect image = Image.Find("some path", attempts:3,region: region);
Mouse.ClickImage(image);
```
Here you can find an Image on the screen providing the path of the image as **PNG** or **JPG** and also the region of the screen we want to look for the image using the **ScreenRegions** class. Notice that region and image are both Rect objects. The **Rect** class is just a custom **Rectangle** struct just to easy manipulation of objects. Images class provide two method to convert from **RectangleToRect** and **RectToRectangle**.

A short way for the above code is
---
```C#
Mouse.ClickImage("some path", attempts: 5,region: region);
```
Manipulate multiple images
---
```C#
Image.FindAll("some path", region: region);
```
The above method works exactly like **Find** except it brings us and **IEnumerable** of **Rect** found in the screen or region provided.

Shortcuts
---
The class shortcuts provide you some easy way to make common windows Shortcuts, but you can also provide your own combination of keys to performs your shortcut.

```C#
KeyBoard.HotKey(KEYCODE.ALT, KEYCODE.F4);
```
Here we performs the close window shortcut, with the method, **HotKey** similar to pyAutoGuid HotKey method. Notice that keys are in **KEYCODE** enum.

Another way to do the same
---
```C#
KeyBoard.Shortcut(Shortcuts.CLOSE_WINDOW);
```
**Shorcut** method is just a wrapper for a **HotKey** overload method that uses  **KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>>** to performs more advance shortcuts like open the task manager, but feel free to use HotKey if you want.

KeyBoard
---
The keyboard class not only provide the methods to **Write** text or **Press** a key, it also provide methods to  copy text and files into the **Clipboard**

```C#
//return true if copied was successful.
KeyBoard.Copy("some text to copy into clipboard");  
KeyBoard.CopyFiles("The file path to copy the file into clipBoard");
KeyBoard.CopyFiles(new List<string>() { "List of file paths to copy in clipboard" });
KeyBoard.Paste(); //Return the text copied in the clipboard.
KeyBoard.PasteFiles("Path to paste the files copied",overrride:True);
```

Notice that the **PasteFiles** method receives an **override** param to override all the files if the path to paste contains them.

Mouse
---
Mouse class Also provide other methods to work with besides of **Click** or **ClickImages** ones.

```C#
//Click the mouse with the right click
Mouse.RightClick();

// Click and image with right click
Mouse.RightClickImage(image);

//Move the mouse in pixel
Mouse.Move(200,300); 

//Move the mouse from the center of and image to a specified direction
Mouse.MoveFromImage(image, Direction.BOTTOM); 

//Return a point X,Y of the current Mouse position
Mouse.Position();

//Performs and scroll with the mouse
Mouse.Scroll(200);
```

Images
---
Images also provide other methods to work with

```C#
//Finds the first occurrence of an image using threads
Image.FindFirst(new List<string>() { "Path1", "Path2" }, attempts: 2, region: ScreenRegions.Complete());

//Find all the images using threads
Image.FindAll(new List<string>() { "Path1", "Path2" }, attempts: 2, region: ScreenRegions.Complete());

//Look for an Image until it Fond it or timeout
Image.FindUntil("Path", region: ScreenRegions);

//Look for and image and the same image further than the first one. 
Image.GetImageAndSubImage("image path");

```
Table images
---
If you have images of the same type and they are sorted as a table, you can look for and specific row and column with: 

```C#
//Look for and image in a row an column specified 
Image.FindRowColumn("image path", rowIndex: 1, colIndex: 1);

///Look for a collection of images in a rowRange and ColumnRange specified
Image.FindRowColumnRange("image path",rowIndexRange: new Range(2,3), colIndexRange: new Range(1,3));
```

Window
---
The class Window is useful if you wants to manipulate Microsoft Windows with basic operations like open, close, maximize etc.

```C#
//Open a window using the specified path
Window.Open("Path of the program to execute", timeToWait: 20);

//Get the active windows if possible
Window.GetActive();

//Those method return windows or collection of windows
Window.GetWindow("window name");       
Window.GetWindowWithParcialName("window name");
Window.GetWindows("Windows name");            
Window.GetWindowWithPartialName("window name");
```

Explorer window
---
If the windows if not a program but an explorer window, **GetWindow** method will return null so we can try with another useful method.

```C#
//Look an explorer window
Window.GetExplorerWindow("window name");

```

Window Object
---
When you get the window as an instance of the class you can get some information about the window and also use the following methods to manipulate it.

Properties
---
```C#
//And identifier for the window
myWindow.Handle;

//All the child elements of the windows if have
myWindow.Children;

//A custom string type of the window
myWindow.WindowType;

//Class name of the windows
myWindow.WindowClassName;

//Title of the windows
myWindow.Title;
```

Methods
---
```C#
//Maximize the current windows
myWindow.Maximize();

//Minimize the current windows
myWindow.Minimize();

//Resize the current windows
myWindow.Resize(200, 300);

//Show the windows if is hide 
myWindow.Show();     

//Hide the current windows
myWindow.Hide();

//Return tru if windows if visible
myWindow.IsVisible();

//Move the current windows
myWindow.Move(300, 400);

//Get the size of the current windows
myWindow.Size();

//Make an ScreenShot of the current window
myWindow.TakeScreenShot("Path to save the ScrenShoot");

//Focus the curren window
myWindow.Focus();

//Close the current windows
myWindow.Close();

//Get the Area as a Rect class of the current window
myWindow.Area();
```

ScreenShot
---
The **ScreenShot** class provide some method to make screenShot

```C#
ScreenShot.Capture("path to save");
ScreenShot.Capture(region,"path to save");
```

## License
This library is totally **Free!.**
