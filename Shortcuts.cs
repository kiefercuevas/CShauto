using System.Collections.Generic;

namespace CShauto
{
    ///<summary>A class that provide some Shortcuts to use with Keyboard Command methods</summary>
    public static class Shortcuts
    {
        ///<summary>Shortcut to perform the window copy Shortcut CTRL + C</summary>
        public static KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>> COPY = new KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>>(new List<KEYCODE>() { KEYCODE.CONTROL }, new List<KEYCODE>() { KEYCODE.C });
        ///<summary>Shortcut to perform the window paste Shortcut CTRL + V</summary>
        public static KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>> PASTE = new KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>>(new List<KEYCODE>() { KEYCODE.CONTROL }, new List<KEYCODE>() { KEYCODE.V });
        ///<summary>Shortcut to perform the window Cut Shortcut CTRL + X</summary>
        public static KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>> CUT = new KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>>(new List<KEYCODE>() { KEYCODE.CONTROL }, new List<KEYCODE>() { KEYCODE.X });
        ///<summary>Shortcut to perform the window undo Shortcut CTRL + Z</summary>
        public static KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>> UNDO = new KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>>(new List<KEYCODE>() { KEYCODE.CONTROL }, new List<KEYCODE>() { KEYCODE.Z });
        ///<summary>Shortcut to perform the window copy Shortcut CTRL + Y</summary>
        public static KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>> REDO = new KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>>(new List<KEYCODE>() { KEYCODE.CONTROL }, new List<KEYCODE>() { KEYCODE.Y });
        ///<summary>Shortcut to perform the window go back Shortcut ALT + LEFT</summary>
        public static KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>> GO_BACK = new KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>>(new List<KEYCODE>() { KEYCODE.ALT }, new List<KEYCODE>() { KEYCODE.LEFT });
        ///<summary>Shortcut to perform the window go forward Shortcut ALT + RIGHT</summary>
        public static KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>> GO_FORWARD = new KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>>(new List<KEYCODE>() { KEYCODE.ALT }, new List<KEYCODE>() { KEYCODE.RIGHT });
        ///<summary>Shortcut to perform the window select all Shortcut CTRL + A</summary>
        public static KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>> SELECT_ALL = new KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>>(new List<KEYCODE>() { KEYCODE.CONTROL }, new List<KEYCODE>() { KEYCODE.A });
        ///<summary>Shortcut to perform the window save Shortcut CTRL + S</summary>
        public static KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>> SAVE = new KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>>(new List<KEYCODE>() { KEYCODE.CONTROL }, new List<KEYCODE>() { KEYCODE.S });
        ///<summary>Shortcut to perform the window delete Shortcut CTRL + D</summary>
        public static KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>> DELETE = new KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>>(new List<KEYCODE>() { KEYCODE.CONTROL }, new List<KEYCODE>() { KEYCODE.D });
        ///<summary>Shortcut to perform the window close Shortcut CTRL + F4</summary>
        public static KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>> CLOSE_WINDOW = new KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>>(new List<KEYCODE>() { KEYCODE.ALT }, new List<KEYCODE>() { KEYCODE.F4 });
        ///<summary>Shortcut to perform the window go down Shortcut CTRL + END</summary>
        public static KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>> GO_DOWN = new KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>>(new List<KEYCODE>() { KEYCODE.CONTROL }, new List<KEYCODE>() { KEYCODE.END });
        ///<summary>Shortcut to perform the window go up Shortcut CTRL + HOME</summary>
        public static KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>> GO_UP = new KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>>(new List<KEYCODE>() { KEYCODE.CONTROL }, new List<KEYCODE>() { KEYCODE.HOME });
        ///<summary>Shortcut to open the top-left menu of window ALT + SPACE</summary>
        public static KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>> OPEN_WINDOW_MENU = new KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>>(new List<KEYCODE>() { KEYCODE.ALT }, new List<KEYCODE>() { KEYCODE.SPACE });
        ///<summary>Shortcut to open the window explorer  LEFT-WIN + E</summary>
        public static KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>> WINDOW_EXPLORER = new KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>>(new List<KEYCODE>() { KEYCODE.L_WIN }, new List<KEYCODE>() { KEYCODE.E });
        ///<summary>Shortcut to focus the type cursor into address search bar of a window ALT + D</summary>
        public static KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>> WINDOW_ADDRESS_BAR = new KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>>(new List<KEYCODE>() { KEYCODE.ALT }, new List<KEYCODE>() { KEYCODE.D });
        ///<summary>Shortcut to open the run command window LEFT-WIN + R</summary>
        public static KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>> WINDOW_RUN = new KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>>(new List<KEYCODE>() { KEYCODE.L_WIN }, new List<KEYCODE>() { KEYCODE.R });
        ///<summary>Shortcut to open the window task manager  CTRL + SHIFT + SCAPE</summary>
        public static KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>> WINDOW_TASK_MANAGER = new KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>>(new List<KEYCODE>() { KEYCODE.CONTROL, KEYCODE.SHIFT }, new List<KEYCODE>() { KEYCODE.ESCAPE });
        ///<summary>Shortcut to switch to other window ALT + TAB + TAB</summary>
        public static KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>> WINDOW_SWITCH = new KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>>(new List<KEYCODE>() { KEYCODE.ALT, KEYCODE.TAB }, new List<KEYCODE>() { KEYCODE.TAB });
        ///<summary>Shortcut to maximize a windows LEFT-WIN + UP</summary>
        public static KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>> WINDOW_MAXIMIZE = new KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>>(new List<KEYCODE>() { KEYCODE.L_WIN }, new List<KEYCODE>() { KEYCODE.UP });
        ///<summary>Shortcut to minimize a windows LEFT-WIN + DOWN</summary>
        public static KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>> WINDOW_MIMINIZE = new KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>>(new List<KEYCODE>() { KEYCODE.L_WIN }, new List<KEYCODE>() { KEYCODE.DOWN });
        ///<summary>Shortcut to minimize all windows LEFT-WIN + M</summary>
        public static KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>> WINDOW_MINIMIZE_ALL = new KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>>(new List<KEYCODE>() { KEYCODE.L_WIN }, new List<KEYCODE>() { KEYCODE.M });
        ///<summary>Shortcut to switch to desktop LEFT-WIN + D</summary>
        public static KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>> WINDOW_SHOW_DESKTOP = new KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>>(new List<KEYCODE>() { KEYCODE.L_WIN }, new List<KEYCODE>() { KEYCODE.D });
        ///<summary>Shortcut to restore all windows if WINDOW_MINIMIZE_ALL was performed SHIFT + LEFT-WIN + D</summary>
        public static KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>> WINDOW_RESTORE_AFTER_MINIMIZE = new KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>>(new List<KEYCODE>() { KEYCODE.SHIFT, KEYCODE.L_WIN }, new List<KEYCODE>() { KEYCODE.M });
        ///<summary>Shortcut to go to task view LEFT-WIN + TAB</summary>
        public static KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>> WINDOW_TASK_VIEW = new KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>>(new List<KEYCODE>() { KEYCODE.L_WIN }, new List<KEYCODE>() { KEYCODE.TAB });
        ///<summary>Shortcut to open the settings window LEFT-WIN + I</summary>
        public static KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>> WINDOW_OPEN_SETTING = new KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>>(new List<KEYCODE>() { KEYCODE.L_WIN }, new List<KEYCODE>() { KEYCODE.I });
        ///<summary>Shortcut to open select menu options if is focussed ALT + DOWN </summary>
        public static KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>> WINDOW_OPEN_SELECT_MENU = new KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>>(new List<KEYCODE>() { KEYCODE.ALT }, new List<KEYCODE>() { KEYCODE.DOWN });
        ///<summary>Shortcut focus the search bar of a windows CTRL + E </summary>
        public static KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>> WINDOW_GO_TO_SEARCH_BAR = new KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>>(new List<KEYCODE>() { KEYCODE.CONTROL }, new List<KEYCODE>() { KEYCODE.E });
        ///<summary>Shortcut to open files or folder properties if it focussed ALT + ENTER </summary>
        public static KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>> WINDOW_SHOW_FILE_PROPERTIES = new KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>>(new List<KEYCODE>() { KEYCODE.ALT }, new List<KEYCODE>() { KEYCODE.ENTER });
        ///<summary>Shortcut to open system information windows WIN + PAUSE </summary>
        public static KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>> WINDOW_SYSTEM_INFORMATION = new KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>>(new List<KEYCODE>() { KEYCODE.L_WIN }, new List<KEYCODE>() { KEYCODE.PAUSE });
        ///<summary>Shortcut to open file menu SHIFT + F10 </summary>
        public static KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>> WINDOW_OPEN_FILE_MENU = new KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>>(new List<KEYCODE>() { KEYCODE.SHIFT }, new List<KEYCODE>() { KEYCODE.F10 });
        ///<summary>Shortcut to delete a file permanently SHIFT + DELETE </summary>
        public static KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>> WINDOW_DELETE_FILE_PERMANENTLY = new KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>>(new List<KEYCODE>() { KEYCODE.SHIFT }, new List<KEYCODE>() { KEYCODE.DELETE });
        ///<summary>Shortcut to show folder files as small CTRL + SHIFT + D4 </summary>
        public static KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>> WINDOW_SHOW_ITEM_SMALL = new KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>>(new List<KEYCODE>() { KEYCODE.CONTROL, KEYCODE.SHIFT }, new List<KEYCODE>() { KEYCODE.D4 });
        ///<summary>Shortcut to show folder files as extra all CTRL + SHIFT + D1 </summary>
        public static KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>> WINDOW_SHOW_ITEM_EXTRA_LARGE = new KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>>(new List<KEYCODE>() { KEYCODE.CONTROL, KEYCODE.SHIFT }, new List<KEYCODE>() { KEYCODE.D1 });
        ///<summary>Shortcut to show folder files as extra large CTRL + SHIFT + D2 </summary>
        public static KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>> WINDOW_SHOW_ITEM_LARGE = new KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>>(new List<KEYCODE>() { KEYCODE.CONTROL, KEYCODE.SHIFT }, new List<KEYCODE>() { KEYCODE.D2 });
        ///<summary>Shortcut to show folder files as medium CTRL + SHIFT + D3 </summary>
        public static KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>> WINDOW_SHOW_ITEM_MEDIUM = new KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>>(new List<KEYCODE>() { KEYCODE.CONTROL, KEYCODE.SHIFT }, new List<KEYCODE>() { KEYCODE.D3 });
        ///<summary>Shortcut to show folder files as list CTRL + SHIFT + D5 </summary>
        public static KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>> WINDOW_SHOW_ITEM_LIST = new KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>>(new List<KEYCODE>() { KEYCODE.CONTROL, KEYCODE.SHIFT }, new List<KEYCODE>() { KEYCODE.D5 });
        ///<summary>Shortcut to show folder files as details CTRL + SHIFT + D6 </summary>
        public static KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>> WINDOW_SHOW_ITEM_DETAILS = new KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>>(new List<KEYCODE>() { KEYCODE.CONTROL, KEYCODE.SHIFT }, new List<KEYCODE>() { KEYCODE.D6 });
        ///<summary>Shortcut to show folder files as tiles CTRL + SHIFT + D7 </summary>
        public static KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>> WINDOW_SHOW_ITEM_TILES = new KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>>(new List<KEYCODE>() { KEYCODE.CONTROL, KEYCODE.SHIFT }, new List<KEYCODE>() { KEYCODE.D7 });
        ///<summary>Shortcut to show folder files as item content CTRL + SHIFT + D8 </summary>
        public static KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>> WINDOW_SHOW_ITEM_CONTENT = new KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>>(new List<KEYCODE>() { KEYCODE.CONTROL, KEYCODE.SHIFT }, new List<KEYCODE>() { KEYCODE.D8 });
        //https://support.microsoft.com/en-us/help/12445/windows-keyboard-shortcuts
        //https://shortcutworld.com/Windows-10-File-Explorer/win/Windows-10-File-Explorer_Shortcuts
    }
}
