using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using WindowsInput;
using WindowsInput.Native;

namespace CShauto
{
    ///<summary>A class that provide some methods to manipulate the keyboard</summary>
    public static class KeyBoard
    {

        ///<summary>Write a secuence of character from the given param</summary>
        /// <param name="word">Secuence of character to write</param>
        public static void Write(string word)
        {
            InputSimulator sim = new InputSimulator();
            sim.Keyboard.TextEntry(word);
        }

        /// <summary>Press a key and amount of quantity and with a given interval and then release it</summary>
        /// <param name="key">Secuence of character to write</param>
        /// <param name="quantity">A</param>
        /// <param name="interval">Interval between press and release </param>
        public static void Press(KEYCODE key, int quantity = 1, double interval = 0)
        {
            InputSimulator sim = new InputSimulator();
            for (int i = 0; i < quantity; i++)
            {
                sim.Keyboard.KeyDown((VirtualKeyCode)key);
                Thread.Sleep(Convert.ToInt32(interval * 1000));
                sim.Keyboard.KeyUp((VirtualKeyCode)key);
            }
        }

        /// <summary>Press a secuence of keys with a given interval and then release them</summary>
        /// <param name="keyToHold">Key to hold</param>
        /// <param name="keyToPres">Key to press while the keyToHold is press, ej: CTRL + K </param>
        public static void HotKey(KEYCODE keyToHold, KEYCODE keyToPres)
        {
            InputSimulator sim = new InputSimulator();
            sim.Keyboard.ModifiedKeyStroke((VirtualKeyCode)keyToHold, (VirtualKeyCode)keyToPres);
        }

        /// <summary>Press a secuence of keys with a given interval and then release them</summary>
        /// <param name="keysToHold">List of keys to hold</param>
        /// <param name="keysToPress">Keys to press while the keyTsoHold are press, ej: CTRL + SHIFT + K + C </param>
        public static void HotKey(IEnumerable<KEYCODE> keysToHold, IEnumerable<KEYCODE> keysToPress)
        {
            InputSimulator sim = new InputSimulator();
            IEnumerable<VirtualKeyCode> newsKeysToHold = keysToHold.Cast<VirtualKeyCode>();
            IEnumerable<VirtualKeyCode> newsKeysToPress = keysToPress.Cast<VirtualKeyCode>();
            sim.Keyboard.ModifiedKeyStroke(newsKeysToHold, newsKeysToPress);
        }

        /// <summary>Hold a secuence of keys and press others keys while holding previus one</summary>
        /// <param name="Shortcuts">List of list of keys to hold and keys to press</param>
        public static void Shortcut(IEnumerable<KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>>> Shortcuts)
        {
            foreach (KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>> keys in Shortcuts)
                HotKey(keys.Key, keys.Value);
        }

        /// <summary>Hold a secuence of keys and press others keys while holding previus one</summary>
        /// <param name="Shortcut">Shortcut to be execute, for more info use class Shortcut</param>
        public static void Shortcut(KeyValuePair<IEnumerable<KEYCODE>, IEnumerable<KEYCODE>> Shortcut)
        {
            HotKey(Shortcut.Key, Shortcut.Value);
        }

        /// <summary>Copy a given text to clickboard</summary>
        /// <param name="text">Text to copy</param>
        public static bool Copy(string text)
        {

            string result = string.Empty;
            Thread newThread = new Thread(() =>
            {
                Clipboard.SetText(text);
                result = Clipboard.GetText();
            });
            newThread.SetApartmentState(ApartmentState.STA);
            newThread.Start();
            newThread.Join();

            return result == text;
        }

        /// <summary>Get a text from the clipboard</summary>
        public static string Paste()
        {
            string result = string.Empty;
            Thread newThread = new Thread(() =>
            {
                result = Clipboard.GetText(TextDataFormat.Text);
            });
            newThread.SetApartmentState(ApartmentState.STA);
            newThread.Start();
            newThread.Join();
            return result;
        }

        /// <summary>Copy a given file to clickboard</summary>
        /// <param name="path">Path of the file to copy</param>
        public static void CopyFiles(string path)
        {
            string result = string.Empty;
            if (!File.Exists(path))
                throw new Exception(string.Format("The file {0} do not exist", path));

            Thread newThread = new Thread(() => Clipboard.SetFileDropList(new StringCollection() { path }));
            newThread.SetApartmentState(ApartmentState.STA);
            newThread.Start();
            newThread.Join();
        }

        /// <summary>Copy a given list of filespath to clickboard</summary>
        /// <param name="paths">Paths of files to copy</param>
        public static void CopyFiles(IEnumerable<string> paths)
        {
            string result = string.Empty;
            StringCollection collection = new StringCollection();
            foreach (string path in paths)
            {
                if (!File.Exists(path))
                    throw new Exception(string.Format("The file {0} do not exist", path));

                collection.Add(path);
            }
            Thread newThread = new Thread(() => Clipboard.SetFileDropList(collection));
            newThread.SetApartmentState(ApartmentState.STA);
            newThread.Start();
            newThread.Join();
        }

        /// <summary>Paste all the copied files in clipboard to an specified path</summary>
        /// <param name="folderPath">Path of the folder to paste the files, *must be a folder*</param>
        /// <param name="overrride">If is true, it will override the files that are the same in the given folder path</param>
        public static void PasteFiles(string folderPath, bool overrride = true)
        {
            if (!Directory.Exists(folderPath))
                throw new Exception("The folder path is not a valid directory");

            StringCollection result = null;
            Thread newThread = new Thread(() => result = Clipboard.GetFileDropList());
            newThread.SetApartmentState(ApartmentState.STA);
            newThread.Start();
            newThread.Join();

            foreach (string clipBoardPath in result)
            {
                string fullPath = Path.Combine(folderPath, Path.GetFileName(clipBoardPath));
                File.Copy(clipBoardPath, fullPath, overrride);
            }

        }

        /// <summary>Get all the files copied in clipboard as filepaths</summary>
        public static IEnumerable<string> GetClipBoardFiles()
        {
            StringCollection result = null;
            ICollection<string> values = new List<string>();
            Thread newThread = new Thread(() => result = Clipboard.GetFileDropList());
            newThread.SetApartmentState(ApartmentState.STA);
            newThread.Start();
            newThread.Join();

            foreach (string clipBoardPath in result)
            {
                values.Add(clipBoardPath);
            }
            return values;
        }

    }
}
