using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace CShauto
{
    ///<summary>A class that provide some methods to manipulate windows including explorer windows</summary>
    public class Window
    {

        ///<summary>Windows handle to identify the current windows</summary>
        public IntPtr Handle { get; private set; }
        ///<summary>Name of the windows</summary>
        public string Title { get; private set; }
        ///<summary>Get the children of a window</summary>
        public ICollection<Window> Children { get; private set; }

        /// <summary>
        /// Get a custom representation of the window class base on https://docs.microsoft.com/en-us/windows/win32/winmsg/about-window-classes documentation
        /// </summary>
        public string WindowType { get; private set; }

        /// <summary>
        /// Get the name of the class that represents the windows
        /// </summary>
        public string WindowClassName { get; private set; }

        private const UInt32 WM_CLOSE = 0x0010;
        private static int Amount = 0;

        /// <summary>Creates a window object with a handle and a window title</summary>
        /// <param name="handle"></param>
        /// <param name="title"></param>
        public Window(IntPtr handle, string title)
        {
            Handle = handle;
            Title = title;

            StringBuilder stringBuilder = new StringBuilder(256);
            GetClassName(handle, stringBuilder, stringBuilder.Capacity);
            WindowType = GetWindowClass(stringBuilder.ToString());
            WindowClassName = stringBuilder.ToString();

            Children = new List<Window>();
            ArrayList handles = new ArrayList();
            EnumedWindow childProc = GetWindowHandle;

            EnumChildWindows(handle, childProc, handles);
            foreach (IntPtr item in handles)
            {
                int capacityChild = GetWindowTextLength(handle) * 2;

                StringBuilder stringBuilderChild = new StringBuilder(capacityChild);
                GetWindowText(handle, stringBuilder, stringBuilderChild.Capacity);

                StringBuilder stringBuilderChild2 = new StringBuilder(256);
                GetClassName(handle, stringBuilderChild2, stringBuilderChild2.Capacity);

                Window win = new Window(item, stringBuilder.ToString());
                win.WindowClassName = stringBuilderChild2.ToString();
                win.WindowType = GetWindowClass(stringBuilderChild2.ToString());
                Children.Add(win);
            }
        }


        private static bool GetWindowHandle(IntPtr windowHandle, ArrayList windowHandles)
        {
            windowHandles.Add(windowHandle);
            return true;
        }

        ///<summary>A class to have better manipulation of windows sizes</summary>
        private struct RectStruct
        {
            public int Left { get; set; }
            public int Top { get; set; }
            public int Right { get; set; }
            public int Bottom { get; set; }
            public int Width { get; set; }
            public int Height { get; set; }

        }

        /// <summary>Open a new Process with the given path and return a window object if the process have a user interface, else return null</summary>
        /// <param name="filePath">Path of the file to look for</param>
        /// <param name="timeToWait">Time to wait until the proccess execute, *Only apply for process with a window interface*</param>
        public static Window Open(string filePath, int timeToWait = -1)
        {
            if (!File.Exists(filePath))
                throw new Exception(string.Format("The filePath {0} is not valid", filePath));

            Process newProcess = Process.Start(filePath);

            if (timeToWait == -1)
                newProcess.WaitForInputIdle();
            else
                newProcess.WaitForInputIdle(timeToWait * 1000);

            if (newProcess != null && newProcess.MainWindowHandle != IntPtr.Zero)
                return new Window(newProcess.MainWindowHandle, newProcess.MainWindowTitle);

            return null;
        }

        /// <summary>Look for the existence of a process with the given name an return the first occurrence of the process as a Window object</summary>
        /// <param name="name">Name of the process</param>
        /// <param name="attempts">Amount of tries that it will look for the window</param>
        /// <param name="waitInterval">Amount ot time it will stop the thread while waiting for the windows in each attemp</param>
        public static Window GetWindow(string name, int attempts = 1, int waitInterval = 1000)
        {
            IEnumerable<Process> currentProcesses = Process.GetProcesses();
            int counter = 0;
            while (counter < attempts)
            {
                foreach (Process p in currentProcesses)
                    if (p.MainWindowHandle != IntPtr.Zero && p.MainWindowTitle == name)
                        return new Window(p.MainWindowHandle, p.MainWindowTitle);

                Helpers.Wait(waitInterval);
                currentProcesses = Process.GetProcesses();
                counter++;
            }
            return null;
        }

        /// <summary>Look for the existence of processes with the given name an return all occurrences of the process as Windows objects, *case sensitive*</summary>
        /// <param name="name">Name of the process</param>
        /// <param name="attempts">Amount of tries that it will look for the window</param>
        /// <param name="waitInterval">Amount ot time it will stop the thread while waiting for the windows in each attemp</param>
        public static IEnumerable<Window> GetWindows(string name, int attempts = 1, int waitInterval = 1000)
        {
            IEnumerable<Process> currentProcesses = Process.GetProcesses();
            ICollection<Window> windows = new List<Window>();
            int counter = 0;
            while (counter < attempts)
            {
                foreach (Process p in currentProcesses)
                    if (p.MainWindowHandle != IntPtr.Zero && p.MainWindowTitle == name)
                        windows.Add(new Window(p.MainWindowHandle, p.MainWindowTitle));

                if (windows.Count > 0)
                    break;

                Helpers.Wait(waitInterval);
                currentProcesses = Process.GetProcesses();
                counter++;
            }
            return windows;
        }

        /// <summary>Look for the existence of a proccess with the given name an return the first ocurrence of the Window as an object 
        /// *This is use for folder and explorer windows, different from files or program execution*</summary>
        /// <param name="name">Name of the window</param>
        /// <param name="exactMatch">if is true, it will look for a window with have the name and not the exactmatch</param>
        /// <param name="attempts">Amount of tries that it will look for the window</param>
        /// <param name="waitInterval">Amount ot time it will stop the thread while waiting for the windows in each attemp</param>
        public static Window GetExplorerWindow(string name, bool exactMatch = false, int attempts = 1, int waitInterval = 1000)
        {
            IntPtr handle = IntPtr.Zero;
            SHDocVw.ShellWindows shellWindows = new SHDocVw.ShellWindows();
            StringBuilder sb = new StringBuilder(256);
            string winTitle = null;
            int counter = 0;

            while (counter < attempts)
            {
                foreach (SHDocVw.InternetExplorer window in shellWindows)
                {
                    handle = new IntPtr(window.HWND);
                    if (handle == IntPtr.Zero)
                        continue;

                    GetWindowText(handle, sb, 256);
                    winTitle = sb.ToString();

                    if (exactMatch)
                        if (winTitle == name)
                            return new Window(handle, winTitle);
                        else
                            if (winTitle.ToLower().Contains(name.ToLower()))
                            return new Window(handle, winTitle);
                }
                Helpers.Wait(waitInterval);
                counter++;
            }
            return null;
        }

        /// <summary>Look for the existense of a process in all processes an return the first ocurrence of a process that contains the given name as a Window object</summary>
        /// <param name="name">Name of the process</param>
        /// <param name="attempts">Amount of tries that it will look for the window</param>
        /// <param name="waitInterval">Amount ot time it will stop the thread while waiting for the windows in each attemp</param>
        public static Window GetWindowWithPartialName(string name, int attempts = 1, int waitInterval = 1000)
        {
            IEnumerable<Process> currentProcesses = Process.GetProcesses();
            int counter = 0;
            while (counter < attempts)
            {
                foreach (Process p in currentProcesses)
                    if (p.MainWindowHandle != IntPtr.Zero && p.MainWindowTitle.ToLower().Contains(name.ToLower()))
                        return new Window(p.MainWindowHandle, p.MainWindowTitle);

                Helpers.Wait(waitInterval);
                currentProcesses = Process.GetProcesses();
                counter++;
            }
            return null;
        }

        /// <summary>Look for the existense of a process in all processes an return the processes that contains the given name as Windows objects</summary>
        /// <param name="name">Name of the process</param>
        /// <param name="attempts">Amount of tries that it will look for at least one window</param>
        /// <param name="waitInterval">Amount ot time it will stop the thread while waiting for the windows in each attemp</param>
        public static IEnumerable<Window> GetWindowsWithPartialName(string name, int attempts = 1, int waitInterval = 1000)
        {
            IEnumerable<Process> currentProcesses = Process.GetProcesses();
            ICollection<Window> windows = new List<Window>();
            int counter = 0;
            while (counter < attempts)
            {
                foreach (Process p in currentProcesses)
                    if (p.MainWindowHandle != IntPtr.Zero && p.MainWindowTitle.ToLower().Contains(name.ToLower()))
                        windows.Add(new Window(p.MainWindowHandle, p.MainWindowTitle));

                if (windows.Count > 0)
                    break;

                Helpers.Wait(waitInterval);
                currentProcesses = Process.GetProcesses();
                counter++;
            }
            return windows;
        }

        /// <summary>
        /// Get the active windows if possible.
        /// </summary>
        /// <returns></returns>
        public static Window GetActive()
        {
            IntPtr handle = GetActiveWindow();
            if (handle != IntPtr.Zero)
            {
                foreach (Process p in Process.GetProcesses())
                    if (p.MainWindowHandle == handle)
                        return new Window(p.MainWindowHandle, p.MainWindowTitle);
            }
            return null;
        }

        /// <summary>Focus the current window</summary>
        public void Focus()
        {
            ShowWindowAsync(this.Handle, (int)ShowWindowCommands.Normal);
            SetForegroundWindow(this.Handle);
        }

        /// <summary>Maximize the current window</summary>
        public bool Maximize()
        {
            return ShowWindowAsync(this.Handle, (int)ShowWindowCommands.Maximize);
        }

        /// <summary>Minimize the current window</summary>
        public bool Minimize()
        {
            return ShowWindowAsync(this.Handle, (int)ShowWindowCommands.Minimize);
        }

        /// <summary>Return the current windows at its first state when the windows was created</summary>
        public bool Restore()
        {
            return ShowWindowAsync(this.Handle, (int)ShowWindowCommands.Restore);
        }

        /// <summary>Return the current windows at its default state</summary>
        public bool DefaultState()
        {
            return ShowWindowAsync(this.Handle, (int)ShowWindowCommands.ShowDefault);
        }

        /// <summary>Hide the current window 
        /// *If the application close with a hide process, this will not be close unless close method 
        /// calls or manually kill the process*</summary>
        public bool Hide()
        {
            return ShowWindowAsync(this.Handle, (int)ShowWindowCommands.Hide);
        }

        /// <summary>Shows the current windows if it was hide</summary>
        public bool Show()
        {
            return ShowWindowAsync(this.Handle, (int)ShowWindowCommands.Show);
        }

        /// <summary>Close the current windows</summary>
        public bool Close()
        {
            return SendMessage(this.Handle, WM_CLOSE, IntPtr.Zero, IntPtr.Zero) == IntPtr.Zero;
        }

        /// <summary>Resize the current window with the given params</summary>
        /// <param name="width">New width of the current windows</param>
        /// <param name="height">New height of the current windows</param>
        public bool Resize(int width, int height)
        {
            return MoveWindow(this.Handle, 0, 0, width, height, true);
        }

        /// <summary>Resize the current window with the given params</summary>
        /// <param name="pixels">this will use as new width and new height of the windows</param>
        public bool Resize(int pixels)
        {
            return MoveWindow(this.Handle, 0, 0, pixels, pixels, true);
        }

        /// <summary>Move the current window with the given params</summary>
        /// <param name="X">New X coordinate of the current windows</param>
        /// <param name="Y">New Y coordinate of the current windows</param>
        public bool Move(int X, int Y)
        {
            RectStruct rect = new RectStruct();
            GetWindowRect(this.Handle, ref rect);

            rect.Width = rect.Right - rect.Left + Amount;
            rect.Height = rect.Bottom - rect.Top + Amount;
            return MoveWindow(this.Handle, X, Y, rect.Width, rect.Height, true);
        }

        /// <summary>Return the position of the windows as X, Y coordinates</summary>
        public Point Position()
        {
            RectStruct rect = new RectStruct();
            GetWindowRect(this.Handle, ref rect);

            rect.Width = rect.Right - rect.Left + Amount;
            rect.Height = rect.Bottom - rect.Top + Amount;
            return new Point(rect.Left, rect.Top);
        }

        /// <summary>Return the Size of the windows as width, height coordinates</summary>
        public Size Size()
        {
            RectStruct rect = new RectStruct();
            GetWindowRect(this.Handle, ref rect);

            rect.Width = rect.Right - rect.Left + Amount;
            rect.Height = rect.Bottom - rect.Top + Amount;
            return new Size(rect.Width, rect.Height);
        }

        /// <summary>Return the position and size of the windows as X, Y, with, height coordinates</summary>
        public Rect Area()
        {
            RectStruct rect = new RectStruct();
            GetWindowRect(this.Handle, ref rect);

            rect.Width = rect.Right - rect.Left + Amount;
            rect.Height = rect.Bottom - rect.Top + Amount;
            return new Rect(rect.Left, rect.Top, rect.Width, rect.Height);
        }

        /// <summary>Check if the current windows is visible</summary>
        public bool IsVisible()
        {
            return IsWindowVisible(this.Handle);
        }

        /// <summary>Make and screenshot of the current windows</summary>
        public bool TakeScreenShot(string fullPath)
        {
            return ScreenShot.Capture(Area(), fullPath);
        }

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        private static extern IntPtr GetActiveWindow();

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, ref RectStruct rectangle);

        [DllImport("kernel32.dll")]
        private static extern int GetProcessId(IntPtr handle);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, IntPtr ProcessId);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetActiveWindow(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetFocus(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern void SwitchToThisWindow(IntPtr hWnd, bool fAltTab);

        [DllImport("user32.dll")]
        private static extern IntPtr GetFocus();

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        private delegate bool EnumedWindow(IntPtr handleWindow, ArrayList handles);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool EnumChildWindows(IntPtr window, EnumedWindow callback, ArrayList lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int GetWindowTextLength(IntPtr hWnd);

        private enum ShowWindowCommands
        {
            /// <summary>
            /// Hides the window and activates another window.
            /// </summary>
            Hide = 0,
            /// <summary>
            /// Activates and displays a window. If the window is minimized or
            /// maximized, the system restores it to its original size and position.
            /// An application should specify this flag when displaying the window
            /// for the first time.
            /// </summary>
            Normal = 1,
            /// <summary>
            /// Activates the window and displays it as a minimized window.
            /// </summary>
            ShowMinimized = 2,
            /// <summary>
            /// Maximizes the specified window.
            /// </summary>
            Maximize = 3, // is this the right value?
            /// <summary>
            /// Activates the window and displays it as a maximized window.
            /// </summary>      
            ShowMaximized = 3,
            /// <summary>
            /// Displays a window in its most recent size and position. This value
            /// is similar to <see cref="Win32.ShowWindowCommand.Normal"/>, except
            /// the window is not activated.
            /// </summary>
            ShowNoActivate = 4,
            /// <summary>
            /// Activates the window and displays it in its current size and position.
            /// </summary>
            Show = 5,
            /// <summary>
            /// Minimizes the specified window and activates the next top-level
            /// window in the Z order.
            /// </summary>
            Minimize = 6,
            /// <summary>
            /// Displays the window as a minimized window. This value is similar to
            /// <see cref="Win32.ShowWindowCommand.ShowMinimized"/>, except the
            /// window is not activated.
            /// </summary>
            ShowMinNoActive = 7,
            /// <summary>
            /// Displays the window in its current size and position. This value is
            /// similar to <see cref="Win32.ShowWindowCommand.Show"/>, except the
            /// window is not activated.
            /// </summary>
            ShowNA = 8,
            /// <summary>
            /// Activates and displays the window. If the window is minimized or
            /// maximized, the system restores it to its original size and position.
            /// An application should specify this flag when restoring a minimized window.
            /// </summary>
            Restore = 9,
            /// <summary>
            /// Sets the show state based on the SW_* value specified in the
            /// STARTUPINFO structure passed to the CreateProcess function by the
            /// program that started the application.
            /// </summary>
            ShowDefault = 10,
            /// <summary>
            ///  <b>Windows 2000/XP:</b> Minimizes a window, even if the thread
            /// that owns the window is not responding. This flag should only be
            /// used when minimizing windows from a different thread.
            /// </summary>
            ForceMinimize = 11
        }
        private string GetWindowClass(string windowClass)
        {
            List<string> values = new List<string>(){
                "ComboLBox","DDEMLEvent","Message","#32768",
                "#32769","#32770","#32771","#32772","Button","Edit","ListBox","MDIClient",
                "ScrollBar","Static",""
            };

            if (windowClass == "#32768")
                return "Menu";
            else if (windowClass == "#32769")
                return "DektopWindow";
            else if (windowClass == "#32770")
                return "DialogBox";
            else if (windowClass == "#32771")
                return "TaskSwitchWindowClass";
            else if (windowClass == "#32772")
                return "IconTitlesClass";
            return values.SingleOrDefault(s => s == windowClass);
        }
    }
}
