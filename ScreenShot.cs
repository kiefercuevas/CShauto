using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

namespace CShauto
{
    ///<summary>A class that provide some methods to make screenShot</summary>
    public static class ScreenShot
    {
        ///<summary>Returns the valid image extensions</summary>
        public static List<string> ValidExtensions = new List<string>() { ".png", ".jpg" };

        ///<summary>Performs an screenShot of the entire screen</summary>
        /// <param name="fullPath">path to save the capture image</param>
        public static bool Capture(string fullPath)
        {
            string folderPath = Path.GetDirectoryName(fullPath);
            string filename = Path.GetFileName(fullPath);

            if (!Directory.Exists(folderPath))
                throw new Exception(string.Format("The Directory {0} is not valid", folderPath));
            if (!ValidExtensions.Contains(Path.GetExtension(filename).ToLower()))
                throw new Exception(string.Format("The file extension is not valid, consider use png or jpg"));

            Rectangle rect = Image.RectToRectangle(ScreenRegions.Complete());
            Bitmap bmp = new Bitmap(rect.Size.Width, rect.Size.Height);
            //string fullPath = Path.Combine(folderPath,filename);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.CopyFromScreen(0, 0, 0, 0, rect.Size);
                bmp.Save(fullPath);
            }
            return File.Exists(fullPath);
        }

        ///<summary>Performs an screenShot of the entire screen</summary>
        /// <param name="region">region as Rect class to take the screenShot</param>
        /// <param name="fullPath">path to save the capture image</param>
        public static bool Capture(Rect region, string fullPath)
        {
            Bitmap bmp = new Bitmap(region.Width, region.Height);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.CopyFromScreen(region.X, region.Y, 0, 0, new Size(region.Width, region.Height));
                bmp.Save(fullPath);
            }
            return File.Exists(fullPath);
        }

        ///<summary>Performs an screenShot of the entire screen</summary>
        /// <param name="region">region as Rectangle struct to take the screenShot</param>
        /// <param name="fullPath">path to save the capture image</param>
        public static bool Capture(Rectangle region, string fullPath)
        {
            Bitmap bmp = new Bitmap(region.Width, region.Height);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.CopyFromScreen(region.X, region.Y, 0, 0, new Size(region.Width, region.Height));
                bmp.Save(fullPath);
            }
            return File.Exists(fullPath);
        }
    }
}
