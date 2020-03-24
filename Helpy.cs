using Quellatalo.Nin.TheEyes.Pattern.CV.Image;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.IO;
using Quellatalo.Nin.TheEyes;
using System.Windows.Forms;
using System.Threading;
using Quellatalo.Nin.TheHands;
using System.Collections.Specialized;
using System.Runtime.InteropServices;
using System.Diagnostics;
using ClosedXML.Excel;
using WindowsInput;
using WindowsInput.Native;
using System.Reflection;
using System.Collections;

namespace Helpy
{
    /// <summary>A class thar provides some methods to locate images on screen</summary>
    public static class Image
    {

        /// <summary>A box with contains the image rect and the name of the image</summary>
        public class Box
        {
            public Rect Rect { get; set; }
            public string ImgName { get; set; }

            public Box() { }

            public Box(Rect rect, string imgName)
            {
                this.Rect = rect;
                this.ImgName = imgName;
            }
        }

        /// <summary>Provide information about the image</summary>
        private class ImageInfo
        {
            public bool Status { get; set; }
            public Rect Region { get; set; }
            public IList<Rect> imgs { get; set; }
            public IList<string> filenames { get; set; }

            public ImageInfo(bool status, Rect region)
            {
                this.Region = region;
                this.Status = status;
            }

            public ImageInfo(bool status, Rect region, bool initList)
            {
                this.Region = region;
                this.Status = status;
                imgs = initList ? new List<Rect>() : null;
                filenames = initList ? new List<string>() : null;
            }
        }

        /// <summary>Locate and image in the screen and return the coordinates (X, Y, Width, height) of the image</summary>
        /// <param name="path">Path of the image to look for</param>
        /// <param name="attempts">Amount of tries to look for the image in the screen</param>
        /// <param name="region">Region of the screen to look for the image, for more details look for ScreenRegions class</param>
        /// <param name="waitInterval">Amount of time in milliseconds to wait for each attempts before starting to look for the image again</param>
        public static Rect Find(string path, int attempts = 0, Rect region = null, int waitInterval = 1000)
        {
            ValidateImage(path);
            for (int i = 0; i <= attempts; i++)
            {
                using (ImagePattern pattern = new ImagePattern(new Bitmap(Bitmap.FromFile(path))))
                {
                    Rectangle currentRegion = region == null ? RectToRectangle(ScreenRegions.Complete()) : RectToRectangle(region);
                    Area area = new Area(currentRegion);

                    Quellatalo.Nin.TheEyes.Match match = area.Find(pattern);
                    if (match == null)
                    {
                        Thread.Sleep(waitInterval);
                        continue;
                    }
                    else
                        return region == null ? RectangleToRect(match.Rectangle) : FixCoordinates(currentRegion, match.Rectangle);
                }
            }
            return null;
        }

        /// <summary>Locate and image in the screen and return the coordinates (X, Y, Width, height) of the image</summary>
        /// <param name="image">image to look for</param>
        /// <param name="attempts">Amount of tries to look for the image in the screen</param>
        /// <param name="region">Region of the screen to look for the image, for more details look for ScreenRegions class</param>
        /// <param name="waitInterval">Amount of time in milliseconds to wait for each attempts before starting to look for the image again</param>
        public static Rect Find(Bitmap image, int attempts = 0, Rect region = null, int waitInterval = 2)
        {
            for (int i = 0; i <= attempts; i++)
            {
                using (ImagePattern pattern = new ImagePattern(new Bitmap(image)))
                {
                    Rectangle currentRegion = region == null ? RectToRectangle(ScreenRegions.Complete()) : RectToRectangle(region);
                    Area area = new Area(currentRegion);

                    Quellatalo.Nin.TheEyes.Match match = area.Find(pattern);
                    if (match == null)
                    {
                        Thread.Sleep(waitInterval);
                        continue;
                    }
                    else
                        return region == null ? RectangleToRect(match.Rectangle) : FixCoordinates(currentRegion, match.Rectangle);
                }
            }
            return null;
        }

        /// <summary>Locate and image in the screen and return true if the image is found</summary>
        /// <param name="path">Path of the image to look for</param>
        /// <param name="attempts">Amount of tries to look for the image in the screen</param>
        /// <param name="region">Region of the screen to look for the image, for more details look for ScreenRegions class</param>
        /// <param name="waitInterval">Amount of time in milliseconds to wait for each attempts before starting to look for the image again</param>
        public static bool Exist(string path, int attempts = 0, Rect region = null, int waitInterval = 1000)
        {
            ValidateImage(path);
            for (int i = 0; i <= attempts; i++)
            {
                using (ImagePattern pattern = new ImagePattern(new Bitmap(Bitmap.FromFile(path))))
                {
                    Rectangle currentRegion = region == null ? RectToRectangle(ScreenRegions.Complete()) : RectToRectangle(region);
                    Area area = new Area(currentRegion);

                    Quellatalo.Nin.TheEyes.Match match = area.Find(pattern);
                    if (match == null)
                    {
                        Thread.Sleep(waitInterval);
                        continue;
                    }
                    else
                        return true;
                }
            }
            return false;
        }

        /// <summary>Locate and image in the screen and return true if the image is found</summary>
        /// <param name="image">image to look for</param>
        /// <param name="attempts">Amount of tries to look for the image in the screen</param>
        /// <param name="region">Region of the screen to look for the image, for more details look for ScreenRegions class</param>
        /// <param name="waitInterval">Amount of time in milliseconds to wait for each attempts before starting to look for the image again</param>
        public static bool Exist(Bitmap image, int attempts = 0, Rect region = null, int waitInterval = 1000)
        {
            for (int i = 0; i <= attempts; i++)
            {
                using (ImagePattern pattern = new ImagePattern(new Bitmap(image)))
                {
                    Rectangle currentRegion = region == null ? RectToRectangle(ScreenRegions.Complete()) : RectToRectangle(region);
                    Area area = new Area(currentRegion);

                    Quellatalo.Nin.TheEyes.Match match = area.Find(pattern);
                    if (match == null)
                    {
                        Thread.Sleep(waitInterval);
                        continue;
                    }
                    else
                        return true;
                }
            }
            return false;
        }

        /// <summary>Locate and image in the screen and amount of time and 
        /// return the coordinates (X, Y, Width, height) of the image,if attempts is 0, it will look for the image until it found</summary>
        /// <param name="path">Path of the image to look for</param>
        /// <param name="region">Region of the screen to look for the image, for more details look for ScreenRegions class</param>
        /// <param name="attempts">Amount of time to look for the image, return null if timeout *if attempts is 0 then it look until it founds the image*</param>
        ///<param name="waitInterval">Amount of time in milliseconds to wait for each attempts before starting to look for the image again</param>
        public static Rect FindUntil(string path, Rect region = null, int attempts = 0, int waitInterval = 1000)
        {
            ValidateImage(path);
            int counter = 0;
            Rect result = null;

            while (result == null)
            {
                result = Find(path, region: region);
                if (attempts != -1)
                {
                    Thread.Sleep(waitInterval);
                    counter += 1;
                    if (counter == attempts)
                        return result;
                }
                else
                {
                    Thread.Sleep(waitInterval);
                }
            }
            return result;
        }

        /// <summary>Locate and image in the screen and amount of time and 
        /// return the coordinates (X, Y, Width, height) of the image,if timeout is 0, it will look for the image until it found</summary>
        /// <param name="image">image to look for</param>
        /// <param name="region">Region of the screen to look for the image, for more details look for ScreenRegions class</param>
        /// <param name="attempts">Amount of time to look for the image, return null if timeout *if attempts is 0 then it look until it founds the image*</param>
        /// <param name="waitInterval">Amount of time in milliseconds to wait for each attempts before starting to look for the image again</param>
        public static Rect FindUntil(Bitmap image, Rect region = null, int attempts = 0, int waitInterval = 1000)
        {
            int counter = 0;
            Rect result = null;

            while (result == null)
            {
                result = Find(image, region: region);
                if (attempts > 0)
                {
                    Thread.Sleep(waitInterval);
                    counter += 1;
                    if (counter == attempts)
                        return result;
                }
                else
                    Thread.Sleep(waitInterval);
            }
            return result;
        }

        /// <summary>Locate all ocurrences of and image in the screen</summary>
        /// <param name="path">Path of the image to look for</param>
        /// <param name="region">Region of the screen to look for the image, for more details look for ScreenRegions class</param>
        public static IEnumerable<Rect> FindAll(string path, Rect region = null)
        {
            ICollection<Rect> results = new List<Rect>();
            ValidateImage(path);
            using (ImagePattern pattern = new ImagePattern(new Bitmap(Bitmap.FromFile(path))))
            {
                Rectangle currentRegion = region == null ? RectToRectangle(ScreenRegions.Complete()) : RectToRectangle(region);
                Area area = new Area(currentRegion);
                IEnumerable<Quellatalo.Nin.TheEyes.Match> matches = area.FindAll(pattern);
                foreach (Quellatalo.Nin.TheEyes.Match match in matches)
                {
                    results.Add(region == null ? RectangleToRect(match.Rectangle) : FixCoordinates(currentRegion, match.Rectangle));
                }
                return results;
            }
        }

        /// <summary>Locate all ocurrences of and image in the screen</summary>
        /// <param name="image">image to look for</param>
        /// <param name="region">Region of the screen to look for the image, for more details look for ScreenRegions class</param>
        public static IEnumerable<Rect> FindAll(Bitmap image, Rect region = null)
        {
            ICollection<Rect> results = new List<Rect>();
            using (ImagePattern pattern = new ImagePattern(image))
            {
                Rectangle currentRegion = region == null ? RectToRectangle(ScreenRegions.Complete()) : RectToRectangle(region);
                Area area = new Area(currentRegion);
                IEnumerable<Quellatalo.Nin.TheEyes.Match> matches = area.FindAll(pattern);
                foreach (Quellatalo.Nin.TheEyes.Match match in matches)
                {
                    results.Add(region == null ? RectangleToRect(match.Rectangle) : FixCoordinates(currentRegion, match.Rectangle));
                }
                return results;
            }
        }

        /// <summary>Look for the first appearance of and image at the given list using threads</summary>
        /// <param name="paths">List of path of the images to look for</param>
        /// <param name="attempts">Amount of time to look for the image, return null if timeout</param>
        /// <param name="region">Region of the screen to look for the image, for more details look for ScreenRegions class</param>
        /// <param name="maxImages">This param is use to limit the number of threads to use, if you have more threads feel free tu increase it</param>
        public static Box FindFirst(IEnumerable<string> paths, int attempts = 0, Rect region = null, int maxImages = 4, int waitInterval = 1000)
        {
            Box currentBox = new Box();
            ImageInfo imgInfo = new ImageInfo(true, region);
            int counter = 0;

            if (paths.Count() > maxImages)
                throw new Exception("The paths quantity cannot be greater than maxImages");

            foreach (string path in paths)
            {
                ValidateImage(path);
                Thread newThread = new Thread(() => ExcecuteActionFindFirst(imgInfo, path, currentBox));
                newThread.IsBackground = true;
                newThread.Start();
            }

            while (imgInfo.Status)
            {
                if (attempts > 0)
                {
                    Thread.Sleep(waitInterval);
                    counter += 1;
                    if (counter == attempts)
                        break;
                }
                else
                    Thread.Sleep(waitInterval);
            }
            return currentBox;
        }

        /// <summary>Look for the first appearance of and image at the given list using threads</summary>
        /// <param name="images">List of key,value pair of images to look for, *key will be returned as image name*</param>
        /// <param name="attempts">Amount of time to look for the image, return null if timeout</param>
        /// <param name="region">Region of the screen to look for the image, for more details look for ScreenRegions class</param>
        /// <param name="maxImages">This param is use to limit the number of threads to use, if you have more threads feel free tu increase it</param>
        /// <param name="waitInterval">Amount of time in milliseconds to wait for each attempts before starting to look for the image again</param>
        public static Box FindFirst(IEnumerable<KeyValuePair<string, Bitmap>> images, int attempts = 0, Rect region = null, int maxImages = 4, int waitInterval = 1000)
        {
            Box currentBox = new Box();
            ImageInfo imgInfo = new ImageInfo(true, region);
            int counter = 0;

            if (images.Count() > maxImages)
                throw new Exception("The paths quantity cannot be greater than maxImages");


            foreach (KeyValuePair<string, Bitmap> image in images)
            {
                Thread newThread = new Thread(() => ExcecuteActionFindFirst(imgInfo, image, currentBox));
                newThread.IsBackground = true;
                newThread.Start();
            }

            while (imgInfo.Status)
            {
                if (attempts > 0)
                {
                    Thread.Sleep(waitInterval);
                    counter += 1;
                    if (counter == attempts)
                        break;
                }
                else
                    Thread.Sleep(waitInterval);
            }
            return currentBox;
        }

        /// <summary>Look for all images at the given list using threads</summary>
        /// <param name="paths">List of path of the images to look for</param>
        /// <param name="attempts">Amount of time to look for the image, return null if timeout</param>
        /// <param name="region">Region of the screen to look for the image, for more details look for ScreenRegions class</param>
        /// <param name="maxImages">This param is use to limit the number of threads to use, if you have more threads feel free tu increase it</param>
        /// <param name="waitInterval">Amount of time in milliseconds to wait for each attempts before starting to look for the image again</param>
        public static IEnumerable<Box> FindAll(IEnumerable<string> paths, int attempts = 0, Rect region = null, int maxImages = 4, int waitInterval = 1000)
        {
            ICollection<Box> currentBoxes = new List<Box>();
            ImageInfo imgInfo = new ImageInfo(true, region, true);
            int counter = 0;

            if (paths.Count() > maxImages)
                throw new Exception("The paths quantity cannot be greater than maxImages");

            foreach (string path in paths)
            {
                ValidateImage(path);
                Thread newThread = new Thread(() => ExcecuteActionFindAll(imgInfo, path));
                newThread.IsBackground = true;
                newThread.Start();
            }

            while (imgInfo.imgs.Count() < paths.Count())
            {
                if (attempts > 0)
                {
                    Thread.Sleep(waitInterval);
                    counter += 1;
                    if (counter == attempts)
                        break;
                }
                else
                    Thread.Sleep(waitInterval);
            }

            for (int i = 0; i < imgInfo.imgs.Count; i++)
            {
                currentBoxes.Add(new Box(imgInfo.imgs[i], imgInfo.filenames[i]));
            }

            return currentBoxes;
        }

        /// <summary>Look for all images at the given list using threads **</summary>
        /// <param name="images">List of key,value pair of images to look for, *key will be returned as image name*</param>
        /// <param name="attempts">Amount of time to look for the image, return null if timeout</param>
        /// <param name="region">Region of the screen to look for the image, for more details look for ScreenRegions class</param>
        /// <param name="maxImages">This param is use to limit the number of threads to use, if you have more threads feel free tu increase it</param>
        /// <param name="waitInterval">Amount of time in milliseconds to wait for each attempts before starting to look for the image again</param>
        public static IEnumerable<Box> FindAll(IEnumerable<KeyValuePair<string, Bitmap>> images, int attempts = 0, Rect region = null, int maxImages = 4, int waitInterval = 1000)
        {
            ICollection<Box> currentBoxes = new List<Box>();
            ImageInfo imgInfo = new ImageInfo(true, region, true);
            int counter = 0;

            if (images.Count() > maxImages)
                throw new Exception("The paths quantity cannot be greater than maxImages");

            foreach (KeyValuePair<string, Bitmap> image in images)
            {
                Thread newThread = new Thread(() => ExcecuteActionFindAll(imgInfo, image));
                newThread.IsBackground = true;
                newThread.Start();
            }

            while (imgInfo.imgs.Count() < images.Count())
            {
                if (attempts > 0)
                {
                    Thread.Sleep(waitInterval);
                    counter += 1;
                    if (counter == attempts)
                        break;
                }
                else
                    Thread.Sleep(waitInterval);
            }

            for (int i = 0; i < imgInfo.imgs.Count; i++)
            {
                currentBoxes.Add(new Box(imgInfo.imgs[i], imgInfo.filenames[i]));
            }

            return currentBoxes;
        }

        /// <summary>Locate all images and organize them as a table, then use index to look for the request image</summary>
        /// <param name="path">Path of the image to look for</param>
        /// <param name="rowIndex">Index to use to look for a row of and image in the table of images found, 
        /// if is -1 or lower it will ignore this index. *it will raise an exception if index if out of range*</param>
        /// <param name="colIndex">Index to use to look for a column of and image in the table of images found, 
        /// if is -1 or lower it will ignore this index. *it will raise an exception if index if out of range*</param>
        public static IEnumerable<Rect> FindRowColumn(string path, int rowIndex = -1, int colIndex = -1)
        {
            IEnumerable<Rect> imgs = FindAll(path);
            List<int> rows = new List<int>();
            List<int> columns = new List<int>();

            if (imgs.Count() == 0)
                return imgs;

            foreach (Rect img in imgs)
            {
                if (!columns.Contains(img.X))
                    columns.Add(img.X);

                if (!rows.Contains(img.Y))
                    rows.Add(img.Y);
            }

            rows.Sort();
            columns.Sort();

            if (rowIndex >= 0 && (rowIndex > rows.Count))
                throw new Exception(string.Format("Index {0} is out of range of founded rows quantity {1} ", rowIndex, rows.Count));

            if (rowIndex >= 0 && (rowIndex > rows.Count))
                throw new Exception(string.Format("Index {0} is out of range of founded columns quantity {1} ", colIndex, columns.Count));

            if (rowIndex >= 0 && colIndex >= 0)
                return imgs.Where((img) => img.X == columns[colIndex] && img.Y == rows[rowIndex]);
            else if (rowIndex >= 0 && colIndex < 0)
                return imgs.Where((img) => img.Y == rows[rowIndex]);
            else if (colIndex >= 0 && rowIndex < 0)
                return imgs.Where((img) => img.X == columns[colIndex]);
            else
                return imgs;
        }

        /// <summary>Locate all images and organize them as a table, then use index to look for the request image</summary>
        /// <param name="image">bitmap of the image to look for</param>
        /// <param name="rowIndex">Index to use to look for a row of and image in the table of images found, 
        /// if is -1 or lower it will ignore this index. *it will raise an exception if index if out of range*</param>
        /// <param name="colIndex">Index to use to look for a column of and image in the table of images found, 
        /// if is -1 or lower it will ignore this index. *it will raise an exception if index if out of range*</param>
        public static IEnumerable<Rect> FindRowColumn(Bitmap image, int rowIndex = -1, int colIndex = -1)
        {
            IEnumerable<Rect> imgs = FindAll(image);
            List<int> rows = new List<int>();
            List<int> columns = new List<int>();

            if (imgs.Count() == 0)
                return imgs;

            foreach (Rect img in imgs)
            {
                if (!columns.Contains(img.X))
                    columns.Add(img.X);

                if (!rows.Contains(img.Y))
                    rows.Add(img.Y);
            }

            rows.Sort();
            columns.Sort();

            if (rowIndex >= 0 && (rowIndex > rows.Count))
                throw new Exception(string.Format("Index {0} is out of range of founded rows quantity {1} ", rowIndex, rows.Count));

            if (rowIndex >= 0 && (rowIndex > rows.Count))
                throw new Exception(string.Format("Index {0} is out of range of founded columns quantity {1} ", colIndex, columns.Count));

            if (rowIndex >= 0 && colIndex >= 0)
                return imgs.Where((img) => img.X == columns[colIndex] && img.Y == rows[rowIndex]);
            else if (rowIndex >= 0 && colIndex < 0)
                return imgs.Where((img) => img.Y == rows[rowIndex]);
            else if (colIndex >= 0 && rowIndex < 0)
                return imgs.Where((img) => img.X == columns[colIndex]);
            else
                return imgs;
        }

        /// <summary>Locate all images and organize them as a table, then use a range of indexes to look for and image</summary>
        /// <param name="path">Path of the image to look for</param>
        /// <param name="rowIndexRange">Range of ndex to use to look for a row of and image in the table of images found, 
        /// if is null it will ignore this index. *it will raise an exception if index if out of range or 
        /// start index if greater or equal to end index*</param>
        /// <param name="colIndexRange">Index to use to look for a column of and image in the table of images found, 
        /// if is null or it will ignore this index. *it will raise an exception if index if out of range or 
        /// start index if greater or equal to end index*</param>
        public static IEnumerable<Rect> FindRowColumnRange(string path, Range rowIndexRange = null, Range colIndexRange = null)
        {
            IEnumerable<Rect> imgs = FindAll(path);
            List<int> rows = new List<int>();
            List<int> columns = new List<int>();

            if (imgs.Count() == 0)
                return imgs;

            foreach (Rect img in imgs)
            {
                if (!columns.Contains(img.X))
                    columns.Add(img.X);

                if (!rows.Contains(img.Y))
                    rows.Add(img.Y);
            }

            rows.Sort();
            columns.Sort();

            ValidateIndex(rowIndexRange, rows.Count, "rows");
            ValidateIndex(colIndexRange, columns.Count, "columns");

            if (rowIndexRange != null && colIndexRange != null)
                return imgs.Where((img) => (img.X > columns[colIndexRange.Start] && img.X < columns[colIndexRange.End]) && (img.Y > rows[rowIndexRange.Start] && img.Y < rows[rowIndexRange.End]));
            else if (rowIndexRange != null && colIndexRange == null)
                return imgs.Where((img) => (img.Y > rows[rowIndexRange.Start] && img.Y < rows[rowIndexRange.End]));
            else if (colIndexRange != null && rowIndexRange == null)
                return imgs.Where((img) => (img.X > columns[colIndexRange.Start] && img.X < columns[colIndexRange.End]));
            else
                return imgs;
        }

        /// <summary>Locate all images and organize them as a table, then use a range of indexes to look for and image</summary>
        /// <param name="image">Bitmap of the image to look for</param>
        /// <param name="rowIndexRange">Range of ndex to use to look for a row of and image in the table of images found, 
        /// if is null it will ignore this index. *it will raise an exception if index if out of range or 
        /// start index if greater or equal to end index*</param>
        /// <param name="colIndexRange">Index to use to look for a column of and image in the table of images found, 
        /// if is null or it will ignore this index. *it will raise an exception if index if out of range or 
        /// start index if greater or equal to end index*</param>
        public static IEnumerable<Rect> FindRowColumnRange(Bitmap image, Range rowIndexRange = null, Range colIndexRange = null)
        {
            IEnumerable<Rect> imgs = FindAll(image);
            List<int> rows = new List<int>();
            List<int> columns = new List<int>();

            if (imgs.Count() == 0)
                return imgs;

            foreach (Rect img in imgs)
            {
                if (!columns.Contains(img.X))
                    columns.Add(img.X);

                if (!rows.Contains(img.Y))
                    rows.Add(img.Y);
            }

            rows.Sort();
            columns.Sort();

            ValidateIndex(rowIndexRange, rows.Count, "rows");
            ValidateIndex(colIndexRange, columns.Count, "columns");

            if (rowIndexRange != null && colIndexRange != null)
                return imgs.Where((img) => (img.X > columns[colIndexRange.Start] && img.X < columns[colIndexRange.End]) && (img.Y > rows[rowIndexRange.Start] && img.Y < rows[rowIndexRange.End])).ToList();
            else if (rowIndexRange != null && colIndexRange == null)
                return imgs.Where((img) => (img.Y > rows[rowIndexRange.Start] && img.Y < rows[rowIndexRange.End])).ToList();
            else if (colIndexRange != null && rowIndexRange == null)
                return imgs.Where((img) => (img.X > columns[colIndexRange.Start] && img.X < columns[colIndexRange.End])).ToList();
            else
                return imgs;
        }

        /// <summary>Look for an Image and a subImage of the given image path 
        /// *A subImage is and image that is further to the left of the screen and it the same image as the given image path*</summary>
        /// <param name="path">Path of the image to look for</param>
        /// <param name="referenceToStart">Image to use as a reference to start to look for the sub Image </param>
        public static IEnumerable<Rect> GetImageAndSubImage(string path, Rect referenceToStart = null)
        {

            IEnumerable<Rect> images = FindAll(path);
            Rect minorImg = null;
            if (images.Count() == 0)
                return null;

            minorImg = referenceToStart == null ? images.First() : referenceToStart;
            foreach (Rect img in images)
            {
                if (img.X > minorImg.X)
                {
                    return new Rect[2] { minorImg, img };
                }
            }
            return new Rect[2] { minorImg, null };
        }

        /// <summary>Look for an Image and a subImage of the given image path 
        /// *A subImage is and image that is further to the left of the screen and it the same image as the given image path*</summary>
        /// <param name="image">Bitmap of the image to look for</param>
        /// <param name="referenceToStart">Image to use as a reference to start to look for the sub Image </param>
        public static IEnumerable<Rect> GetImageAndSubImage(Bitmap image, Rect referenceToStart = null)
        {
            IEnumerable<Rect> images = FindAll(image);
            Rect minorImg = null;
            if (images.Count() == 0)
                return null;

            minorImg = referenceToStart == null ? images.First() : referenceToStart;
            foreach (Rect img in images)
            {
                if (img.X > minorImg.X)
                    return new Rect[2] { minorImg, img };
            }
            return new Rect[2] { minorImg, null };
        }

        /// <summary>Convert a Rect object to a Rectangle struct</summary>
        /// <param name="rect">Rect object to convert</param>
        public static Rectangle RectToRectangle(Rect rect)
        {
            return new Rectangle(rect.X, rect.Y, rect.Width, rect.Height);
        }

        /// <summary>Convert a Rectangle struct to a Rect object</summary>
        /// <param name="rectangle">Rectangle object to convert</param>
        public static Rect RectangleToRect(Rectangle rectangle)
        {
            return new Rect(rectangle.X, rectangle.Y, rectangle.Width, rectangle.Height);
        }

        //private methods
        private static void ValidateIndex(Range range, int quantity, string type)
        {
            if (range != null && (range.Start > quantity))
                throw new Exception(string.Format("Index {0} is out of range of founded {1} quantity {2} ", range, type, quantity));
            else if (range != null && range.Start >= range.End)
                throw new Exception("Start value cannot be greater or equal than End value");
            else if (range != null && range.Start < 0)
                throw new Exception("Start value cannot be lower than 0");
            else if (range != null && range.End < 0)
                throw new Exception("End value cannot be lower than 0");
        }

        private static void ExcecuteActionFindFirst(ImageInfo info, string path, Box box)
        {
            Rect rect = null;
            while (info.Status)
            {
                rect = Find(path, region: info.Region);

                if (rect != null)
                {
                    info.Status = false;
                    box.Rect = rect;
                    box.ImgName = path;
                    break;
                }
            }
        }

        private static void ExcecuteActionFindFirst(ImageInfo info, KeyValuePair<string, Bitmap> image, Box box)
        {
            Rect rect = null;
            while (info.Status)
            {
                rect = Find(image.Value, region: info.Region);

                if (rect != null)
                {
                    info.Status = false;
                    box.Rect = rect;
                    box.ImgName = image.Key;
                    break;
                }
            }
        }

        private static void ExcecuteActionFindAll(ImageInfo info, string path)
        {
            Rect rect = null;
            while (info.Status)
            {
                rect = Find(path, region: info.Region);

                if (rect != null)
                {
                    info.imgs.Add(rect);
                    info.filenames.Add(path);
                    break;
                }
            }
        }

        private static void ExcecuteActionFindAll(ImageInfo info, KeyValuePair<string, Bitmap> image)
        {
            Rect rect = null;
            while (info.Status)
            {
                rect = Find(image.Value, region: info.Region);

                if (rect != null)
                {
                    info.imgs.Add(rect);
                    info.filenames.Add(image.Key);
                    break;
                }
            }
        }

        /// <summary>
        /// HightLight a given region with the specified color
        /// </summary>
        /// <param name="rectToHightLight">The region to hightLight</param>
        /// <param name="hightLightColor">The color to use to fill the region</param>
        private static void HighLight(Rect rectToHightLight, Color hightLightColor)
        {
            Rectangle currentRegion = rectToHightLight == null ? RectToRectangle(ScreenRegions.Complete()) : RectToRectangle(rectToHightLight);
            Area area = new Area(currentRegion);
            area.Highlight(new SolidBrush(hightLightColor));
        }

        /// <summary>
        /// HightLight a given region with the specified color and opacity
        /// </summary>
        /// <param name="rectToHightLight">The region to hightLight</param>
        /// <param name="hightLightColor">The color to use to fill the region</param>
        /// <param name="opacity">The opacity of the color</param>
        public static void HighLight(Rect rectToHightLight, Color hightLightColor, int opacity)
        {
            Rectangle currentRegion = rectToHightLight == null ? RectToRectangle(ScreenRegions.Complete()) : RectToRectangle(rectToHightLight);
            Area area = new Area(currentRegion);
            area.Highlight(new SolidBrush(Color.FromArgb(opacity, hightLightColor)));
        }

        /// <summary>
        /// HightLight a given region with the specified brush
        /// </summary>
        /// <param name="rectToHightLight">The region to hightLight</param>
        /// <param name="brush">The brush to use to hightLight the given region</param>
        public static void HighLight(Rect rectToHightLight, Brush brush)
        {
            Rectangle currentRegion = rectToHightLight == null ? RectToRectangle(ScreenRegions.Complete()) : RectToRectangle(rectToHightLight);
            Area area = new Area(currentRegion);
            area.Highlight(brush);
        }

        /// <summary>
        /// HightLight a given region with the specified pen
        /// </summary>
        /// <param name="rectToHightLight">The region to hightLight</param>
        /// <param name="pen">The pen to use to hightLight the given region</param>
        public static void HighLight(Rect rectToHightLight, Pen pen)
        {
            Rectangle currentRegion = rectToHightLight == null ? RectToRectangle(ScreenRegions.Complete()) : RectToRectangle(rectToHightLight);
            Area area = new Area(currentRegion);
            area.Highlight(pen);
        }

        /// <summary>
        /// Remove all hightLight from current screen
        /// </summary>
        public static void RemoveHighLight()
        {
            Area.ClearHighlight();
        }

        private static Rect FixCoordinates(Rectangle firstRect, Rectangle secondRect)
        {
            return new Rect(firstRect.X + secondRect.X, firstRect.Y + secondRect.Y, secondRect.Width, secondRect.Height);
        }
        /// <summary>A helper to validate an image path</summary>
        private static void ValidateImage(string path)
        {
            if (!File.Exists(path))
                throw new Exception(string.Format("The path {} is not valid", path));
        }

    }

    /// <summary>A class to get screen regions</summary>
    public static class ScreenRegions
    {

        /// <summary>Return the complete coordinate(X,Y, width, height) of the screen</summary>
        public static Rect Complete()
        {
            return Image.RectangleToRect(Screen.PrimaryScreen.Bounds);
        }

        /// <summary>Return the Top coordinate(X,Y, width, height) of the screen</summary>
        public static Rect Top()
        {
            Rectangle screen = Screen.PrimaryScreen.Bounds;
            return Image.RectangleToRect(new Rectangle(0, 0, screen.Width, screen.Height / 2));
        }

        /// <summary>Return the Bottom coordinate(X,Y, width, height) of the screen</summary>
        public static Rect Bottom()
        {
            Rectangle screen = Screen.PrimaryScreen.Bounds;
            return Image.RectangleToRect(new Rectangle(0, screen.Height / 2, screen.Width, screen.Height / 2));
        }

        /// <summary>Return the left coordinate(X,Y, width, height) of the screen</summary>
        public static Rect Left()
        {
            Rectangle screen = Screen.PrimaryScreen.Bounds;
            return Image.RectangleToRect(new Rectangle(0, 0, screen.Width / 2, screen.Height));
        }

        /// <summary>Return the right coordinate(X,Y, width, height) of the screen</summary>
        public static Rect Right()
        {
            Rectangle screen = Screen.PrimaryScreen.Bounds;
            return Image.RectangleToRect(new Rectangle(screen.Width / 2, 0, screen.Width / 2, screen.Height));
        }

        /// <summary>Return the top left coordinate(X,Y, width, height) of the screen</summary>
        public static Rect TopLeft()
        {
            Rectangle screen = Screen.PrimaryScreen.Bounds;
            return Image.RectangleToRect(new Rectangle(0, 0, screen.Width / 2, screen.Height / 2));
        }

        /// <summary>Return the top right coordinate(X,Y, width, height) of the screen</summary>
        public static Rect TopRight()
        {
            Rectangle screen = Screen.PrimaryScreen.Bounds;
            return Image.RectangleToRect(new Rectangle(screen.Width / 2, 0, screen.Width / 2, screen.Height / 2));
        }

        /// <summary>Return the bottom left coordinate(X,Y, width, height) of the screen</summary>
        public static Rect BottomLeft()
        {
            Rectangle screen = Screen.PrimaryScreen.Bounds;
            return Image.RectangleToRect(new Rectangle(0, screen.Height / 2, screen.Width / 2, screen.Height / 2));
        }

        /// <summary>Return the bottom right coordinate(X,Y, width, height) of the screen</summary>
        public static Rect BottomRight()
        {
            Rectangle screen = Screen.PrimaryScreen.Bounds;
            return Image.RectangleToRect(new Rectangle(screen.Width / 2, screen.Height / 2, screen.Width / 2, screen.Height / 2));
        }

    }

    /// <summary>A class that provide some methods to manipulate the mouse</summary>
    public static class Mouse
    {
        /// <summary>Performs a click in a given point in the screen</summary>
        /// <param name="point">The point to perform the click</param>
        /// <param name="numberOfClicks">Amount of click to perform</param>
        public static void Click(Point point, int numberOfClicks = 1)
        {
            point = FixPoint(point);
            MouseHandler handler = new MouseHandler();

            for (int i = 0; i < numberOfClicks; i++)
            {
                handler.Click(point);
            }
        }

        /// <summary>Performs a click in the curren position of the mouse</summary>
        /// <param name="numberOfClicks">Amount of click to perform</param>
        public static void Click(int numberOfClicks = 1)
        {
            MouseHandler handler = new MouseHandler();
            for (int i = 0; i < numberOfClicks; i++)
            {
                handler.Click();
            }
        }

        /// <summary>Performs a right click in a given point in the screen</summary>
        /// <param name="point">The point to perform the click</param>
        /// <param name="numberOfClicks">Amount of click to perform</param>
        public static void RightClick(Point point, int numberOfClicks = 1)
        {
            point = FixPoint(point);
            MouseHandler handler = new MouseHandler();

            for (int i = 0; i < numberOfClicks; i++)
            {
                handler.RightClick(point);
            }
        }

        /// <summary>Performs a right click in the curren position of the mouse</summary>
        /// <param name="numberOfClicks">Amount of click to perform</param>
        public static void RightClick(int numberOfClicks = 1)
        {
            MouseHandler handler = new MouseHandler();
            for (int i = 0; i < numberOfClicks; i++)
            {
                handler.RightClick();
            }

        }

        /// <summary>Looks for the given path of the image and performs a click in the center of the image</summary>
        /// <param name="path">Path of the image</param>
        /// <param name="numberOfClicks">Amount of click to perform</param>
        /// <param name="region">Region of the screen to look for the image, for more details look for ScreenRegions class</param>
        /// <param name="attempts">Amount of tries to look for the image in the screen</param>
        public static Rect ClickImage(string path, int numberOfClicks = 1, Rect region = null, int attempts = 1)
        {
            Rect image = Image.Find(path, region: region, attempts: attempts);
            MouseHandler handler = new MouseHandler();

            if (image == null)
                return image;

            for (int i = 0; i < numberOfClicks; i++)
            {
                handler.Click(Center(image));
            }
            return image;
        }

        /// <summary>Performs a click in the center of the rect image object</summary>
        /// <param name="image">Rect image object to click</param>
        /// <param name="numberOfClicks">Amount of click to perform</param>
        public static void ClickImage(Rect image, int numberOfClicks = 1)
        {
            MouseHandler handler = new MouseHandler();
            for (int i = 0; i < numberOfClicks; i++)
            {
                handler.Click(Center(image));
            }
        }

        /// <summary>Looks for the given path of the image and performs a right click in the center of the image</summary>
        /// <param name="path">Path of the image</param>
        /// <param name="numberOfClicks">Amount of click to perform</param>
        /// <param name="region">Region of the screen to look for the image, for more details look for ScreenRegions class</param>
        public static Rect RightClickImage(string path, int numberOfClicks = 1, Rect region = null)
        {
            Rect image = Image.Find(path, region: region);
            MouseHandler handler = new MouseHandler();

            if (image == null)
                return image;

            for (int i = 0; i < numberOfClicks; i++)
            {
                handler.RightClick(Center(image));
            }
            return image;
        }

        /// <summary>Performs a right click in the center of the rect image object</summary>
        /// <param name="image">Rect image object to click</param>
        /// <param name="numberOfClicks">Amount of click to perform</param>
        public static void RightClickImage(Rect image, int numberOfClicks = 1)
        {
            MouseHandler handler = new MouseHandler();
            for (int i = 0; i < numberOfClicks; i++)
            {
                handler.RightClick(Center(image));
            }
        }

        /// <summary>Return the curren position of the mouse in the screen
        /// </summary>
        public static Point Position()
        {
            MouseHandler handler = new MouseHandler();
            return handler.GetPosition();
        }

        /// <summary>Move the mouse to a given point in the screen</summary>
        /// <param name="point">Point in the screen to move</param>
        /// <param name="isRelative">If isRelativeis set to True, it will move the cursor starting from its current positions and not 
        /// from the beginning of the screen </param>
        public static bool Move(Point point, bool isRelative = false)
        {
            MouseHandler handler = new MouseHandler();
            Point currenPosition = Position();
            if (isRelative)
                handler.Move(point);
            else
                handler.MoveTo(new List<Point>() { currenPosition, point });

            Point lastPosition = Position();
            if (currenPosition.X == lastPosition.X && currenPosition.Y == lastPosition.Y)
                return false;

            return true;
        }

        /// <summary>Move the mouse to a given point in the screen</summary>
        /// <param name="x">Horizontal coordinate to move</param>
        /// <param name="y">Vertical coordinate to move</param>
        /// <param name="isRelative">If isRelativeis set to True, it will move the cursor starting from its current positions and not 
        /// from the beginning of the screen </param>
        public static bool Move(int x, int y, bool isRelative = false)
        {
            MouseHandler handler = new MouseHandler();
            Point currenPosition = Position();
            if (isRelative)
                handler.Move(x, y);
            else
                handler.MoveTo(new List<Point>() { currenPosition, new Point(x, y) });

            Point lastPosition = Position();
            if (currenPosition.X == lastPosition.X && currenPosition.Y == lastPosition.Y)
                return false;

            return true;
        }

        /// <summary>Move the mouse to the center of an image in the screen</summary>
        /// <param name="image">rect image object</param>
        /// <param name="isRelative">If isRelativeis set to True, it will move the cursor starting from its current positions and not 
        /// from the beginning of the screen </param>
        public static bool Move(Rect image, bool isRelative = false)
        {
            MouseHandler handler = new MouseHandler();
            Point currenPosition = Position();
            if (isRelative)
                handler.Move(Center(image));
            else
                handler.MoveTo(new List<Point>() { currenPosition, Center(image) });

            Point lastPosition = Position();
            if (currenPosition.X == lastPosition.X && currenPosition.Y == lastPosition.Y)
                return false;

            return true;
        }

        /// <summary>Look for and image at the given path and move the cursor at the center of the image</summary>
        /// <param name="path">Path of the image</param>
        /// <param name="region">Region of the screen to look for the image, for more details look for ScreenRegions class</param>
        /// <param name="isRelative">If isRelativeis set to True, it will move the cursor starting from its current positions and not 
        /// from the beginning of the screen </param>
        public static bool Move(string path, Rect region, bool isRelative = false)
        {
            MouseHandler handler = new MouseHandler();
            Rect image = Image.Find(path, region: region);

            if (image == null)
                return false;

            Point currenPosition = Position();
            if (isRelative)
                handler.Move(Center(image));
            else
                handler.MoveTo(new List<Point>() { currenPosition, Center(image) });

            Point lastPosition = Position();
            if (currenPosition.X == lastPosition.X && currenPosition.Y == lastPosition.Y)
                return false;

            return true;
        }

        /// <summary>Look for and image at the given path and move the cursor at the center of the image and then move the cursor and
        /// amount of pixel starting from the center of the image</summary>
        /// <param name="path">Path of the image</param>
        /// <param name="increment">Amount of pixel to move relative to the image position</param>
        /// <param name="direction">Direction to move the cursor starting from the center of the image</param>
        /// <param name="region">Region of the screen to look for the image, for more details look for ScreenRegions class</param>
        public static Point MoveFromImage(string path, int increment, Direction direction, Rect region = null)
        {
            switch (direction)
            {
                case Direction.BOTTOM: return MoveBottomFromImage(path, increment, region);
                case Direction.TOP: return MoveTopFromImage(path, increment, region);
                case Direction.LEFT: return MoveLeftFromImage(path, increment, region);
                case Direction.RIGHT: return MoveRightFromImage(path, increment, region);
                case Direction.LEFTUP: return MoveLeftUpFromImage(path, increment, region);
                case Direction.LEFTDOWN: return MoveLeftDownFromImage(path, increment, region);
                case Direction.RIGHTUP: return MoveRightUpFromImage(path, increment, region);
                case Direction.RIGHTDOWN: return MoveRightUpFromImage(path, increment, region);
                default: return new Point(-1, -1);
            }
        }

        /// <summary>Move the cursor at the center of the image then move the cursor and amount of pixel 
        /// starting from the center of the image</summary>
        /// <param name="path">Path of the image</param>
        /// <param name="increment">Amount of pixel to move relative to the image position</param>
        /// <param name="direction">Direction to move the cursor starting from the center of the image</param>
        public static Point MoveFromImage(Rect image, int increment, Direction direction)
        {
            switch (direction)
            {
                case Direction.BOTTOM: return MoveBottomFromImage(image, increment);
                case Direction.TOP: return MoveTopFromImage(image, increment);
                case Direction.LEFT: return MoveLeftFromImage(image, increment);
                case Direction.RIGHT: return MoveRightFromImage(image, increment);
                case Direction.LEFTUP: return MoveLeftUpFromImage(image, increment);
                case Direction.LEFTDOWN: return MoveLeftDownFromImage(image, increment);
                case Direction.RIGHTUP: return MoveRightUpFromImage(image, increment);
                case Direction.RIGHTDOWN: return MoveRightUpFromImage(image, increment);
                default: return new Point(-1, -1);
            }
        }

        /// <summary>Look for and image at the given path and move the cursor at the center of the image and 
        /// then move the cursor and amount of pixel to the right of the center of the image</summary>
        /// <param name="path">Path of the image</param>
        /// <param name="increment">Amount of pixel to move relative to the image position</param>
        /// <param name="region">Region of the screen to look for the image, for more details look for ScreenRegions class</param>
        public static Point MoveRightFromImage(string path, int increment, Rect region = null)
        {
            Rect image = Image.Find(path, region: region);
            MouseHandler handler = new MouseHandler();

            if (image == null)
                return new Point(-1, -1);

            Point point = Center(image);
            Move(point);

            Point operations = new Point((image.Width / 2) + increment, 0);
            handler.Move(operations);
            return Position();
        }

        /// <summary>Move the cursor at the center of the image and then move the cursor and amount of pixel to the right 
        /// of the center of the image</summary>
        /// <param name="image">Image rect object to use</param>
        /// <param name="increment">Amount of pixel to move relative to the image position</param>
        public static Point MoveRightFromImage(Rect image, int increment)
        {
            MouseHandler handler = new MouseHandler();
            Point point = Center(image);
            Move(point);

            Point operations = new Point((image.Width / 2) + increment, 0);
            handler.Move(operations);
            return Position();
        }

        /// <summary>Look for and image at the given path and move the cursor at the center of the image and 
        /// then move the cursor and amount of pixel to the right up of the center of the image</summary>
        /// <param name="path">Path of the image</param>
        /// <param name="increment">Amount of pixel to move relative to the image position</param>
        /// <param name="region">Region of the screen to look for the image, for more details look for ScreenRegions class</param>
        public static Point MoveRightUpFromImage(string path, int increment, Rect region = null)
        {
            Rect image = Image.Find(path, region: region);
            MouseHandler handler = new MouseHandler();

            if (image == null)
                return new Point(-1, -1);

            Point point = Center(image);
            Move(point);

            Point operations = new Point((image.Width / 2) + increment, (-1 * (image.Height / 2) - increment));
            handler.Move(operations);
            return Position();
        }

        /// <summary>Move the cursor at the center of the image and then move the cursor and amount of pixel to the right up
        /// of the center of the image</summary>
        /// <param name="image">Image rect object to use</param>
        /// <param name="increment">Amount of pixel to move relative to the image position</param>
        public static Point MoveRightUpFromImage(Rect image, int increment)
        {
            MouseHandler handler = new MouseHandler();
            Point point = Center(image);
            Move(point);

            Point operations = new Point((image.Width / 2) + increment, (-1 * (image.Height / 2) - increment));
            handler.Move(operations);
            return Position();
        }

        /// <summary>Look for and image at the given path and move the cursor at the center of the image and 
        /// then move the cursor and amount of pixel to the right down of the center of the image</summary>
        /// <param name="path">Path of the image</param>
        /// <param name="increment">Amount of pixel to move relative to the image position</param>
        /// <param name="region">Region of the screen to look for the image, for more details look for ScreenRegions class</param>
        public static Point MoveRightDownFromImage(string path, int increment, Rect region = null)
        {
            Rect image = Image.Find(path, region: region);
            MouseHandler handler = new MouseHandler();

            if (image == null)
                return new Point(-1, -1);

            Point point = Center(image);
            Move(point);

            Point operations = new Point((image.Width / 2) + increment, ((image.Height / 2) + increment));
            handler.Move(operations);
            return Position();
        }

        /// <summary>Move the cursor at the center of the image and then move the cursor and amount of pixel to the right down
        /// of the center of the image</summary>
        /// <param name="image">Image rect object to use</param>
        /// <param name="increment">Amount of pixel to move relative to the image position</param>
        public static Point MoveRightDownFromImage(Rect image, int increment)
        {
            MouseHandler handler = new MouseHandler();
            Point point = Center(image);
            Move(point);

            Point operations = new Point((image.Width / 2) + increment, ((image.Height / 2) + increment));
            handler.Move(operations);
            return Position();
        }

        /// <summary>Look for and image at the given path and move the cursor at the center of the image and 
        /// then move the cursor and amount of pixel to the left of the center of the image</summary>
        /// <param name="path">Path of the image</param>
        /// <param name="increment">Amount of pixel to move relative to the image position</param>
        /// <param name="region">Region of the screen to look for the image, for more details look for ScreenRegions class</param>
        public static Point MoveLeftFromImage(string path, int increment, Rect region = null)
        {

            Rect image = Image.Find(path, region: region);
            MouseHandler handler = new MouseHandler();

            if (image == null)
                return new Point(-1, -1);

            Point point = Center(image);
            Move(point);

            Point operations = new Point(-1 * ((image.Width / 2) + increment), 0);
            handler.Move(operations);
            return Position();
        }

        /// <summary>Move the cursor at the center of the image and then move the cursor and amount of pixel to the left
        /// of the center of the image</summary>
        /// <param name="image">Image rect object to use</param>
        /// <param name="increment">Amount of pixel to move relative to the image position</param>
        public static Point MoveLeftFromImage(Rect image, int increment)
        {
            MouseHandler handler = new MouseHandler();
            Point point = Center(image);
            Move(point);

            Point operations = new Point(-1 * ((image.Width / 2) + increment), 0);
            handler.Move(operations);
            return Position();
        }

        /// <summary>Look for and image at the given path and move the cursor at the center of the image and 
        /// then move the cursor and amount of pixel to the left up of the center of the image</summary>
        /// <param name="path">Path of the image</param>
        /// <param name="increment">Amount of pixel to move relative to the image position</param>
        /// <param name="region">Region of the screen to look for the image, for more details look for ScreenRegions class</param>
        public static Point MoveLeftUpFromImage(string path, int increment, Rect region = null)
        {
            Rect image = Image.Find(path, region: region);
            MouseHandler handler = new MouseHandler();

            if (image == null)
                return new Point(-1, -1);

            Point point = Center(image);
            Move(point);

            Point operations = new Point(-1 * ((image.Width / 2) + increment), (-1 * (image.Height / 2) - increment));
            handler.Move(operations);
            return Position();
        }

        /// <summary>Move the cursor at the center of the image and then move the cursor and amount of pixel to the left up 
        /// of the center of the image</summary>
        /// <param name="image">Image rect object to use</param>
        /// <param name="increment">Amount of pixel to move relative to the image position</param>
        public static Point MoveLeftUpFromImage(Rect image, int increment)
        {
            MouseHandler handler = new MouseHandler();
            Point point = Center(image);
            Move(point);

            Point operations = new Point(-1 * ((image.Width / 2) + increment), (-1 * (image.Height / 2) - increment));
            handler.Move(operations);
            return Position();
        }

        /// <summary>Validates if the point is valid</summary>
        /// <param name="point">Point to validate</param>
        public static bool IsValidPoint(Point point)
        {
            if (point.X == -1)
                return false;
            if (point.Y == -1)
                return false;

            return true;
        }

        /// <summary>Look for and image at the given path and move the cursor at the center of the image and 
        /// then move the cursor and amount of pixel to the left down of the center of the image</summary>
        /// <param name="path">Path of the image</param>
        /// <param name="increment">Amount of pixel to move relative to the image position</param>
        /// <param name="region">Region of the screen to look for the image, for more details look for ScreenRegions class</param>
        public static Point MoveLeftDownFromImage(string path, int increment, Rect region = null)
        {
            Rect image = Image.Find(path, region: region);
            MouseHandler handler = new MouseHandler();

            if (image == null)
                return new Point(-1, -1);

            Point point = Center(image);
            Move(point);

            Point operations = new Point((-1 * ((image.Width / 2)) - increment), (image.Height / 2) + increment);
            handler.Move(operations);
            return Position();
        }

        /// <summary>Move the cursor at the center of the image and then move the cursor and amount of pixel to the left down 
        /// of the center of the image</summary>
        /// <param name="image">Image rect object to use</param>
        /// <param name="increment">Amount of pixel to move relative to the image position</param>
        public static Point MoveLeftDownFromImage(Rect image, int increment)
        {
            MouseHandler handler = new MouseHandler();
            Point point = Center(image);
            Move(point);

            Point operations = new Point((-1 * ((image.Width / 2)) - increment), (image.Height / 2) + increment);
            handler.Move(operations);
            return Position();
        }

        /// <summary>Look for and image at the given path and move the cursor at the center of the image and 
        /// then move the cursor and amount of pixel to the top of the center of the image</summary>
        /// <param name="path">Path of the image</param>
        /// <param name="increment">Amount of pixel to move relative to the image position</param>
        /// <param name="region">Region of the screen to look for the image, for more details look for ScreenRegions class</param>
        public static Point MoveTopFromImage(string path, int increment, Rect region = null)
        {
            Rect image = Image.Find(path, region: region);
            MouseHandler handler = new MouseHandler();

            if (image == null)
                return new Point(-1, -1);

            Point point = Center(image);
            Move(point);

            Point operations = new Point(0, (-1 * (image.Height / 2)) - increment);
            handler.Move(operations);
            return Position();
        }

        /// <summary>Move the cursor at the center of the image and then move the cursor and amount of pixel to the top 
        /// of the center of the image</summary>
        /// <param name="image">Image rect object to use</param>
        /// <param name="increment">Amount of pixel to move relative to the image position</param>
        public static Point MoveTopFromImage(Rect image, int increment)
        {
            MouseHandler handler = new MouseHandler();
            Point point = Center(image);
            Move(point);

            Point operations = new Point(0, (-1 * (image.Height / 2)) - increment);
            handler.Move(operations);
            return Position();
        }

        /// <summary>Look for and image at the given path and move the cursor at the center of the image and 
        /// then move the cursor and amount of pixel to the bottom of the center of the image</summary>
        /// <param name="path">Path of the image</param>
        /// <param name="increment">Amount of pixel to move relative to the image position</param>
        /// <param name="region">Region of the screen to look for the image, for more details look for ScreenRegions class</param>
        public static Point MoveBottomFromImage(string path, int increment, Rect region = null)
        {
            Rect image = Image.Find(path, region: region);
            MouseHandler handler = new MouseHandler();

            if (image == null)
                return new Point(-1, -1);

            Point point = Center(image);
            Move(point);

            Point operations = new Point(0, (image.Height / 2) + increment);
            handler.Move(operations);
            return Position();
        }

        /// <summary>Move the cursor at the center of the image and then move the cursor and amount of pixel to the bottom 
        /// of the center of the image</summary>
        /// <param name="image">Image rect object to use</param>
        /// <param name="increment">Amount of pixel to move relative to the image position</param>
        public static Point MoveBottomFromImage(Rect image, int increment)
        {
            MouseHandler handler = new MouseHandler();
            Point point = Center(image);
            Move(point);

            Point operations = new Point(0, (image.Height / 2) + increment);
            handler.Move(operations);
            return Position();
        }

        /// <summary>
        /// Move the cursor at the center of the image
        /// </summary>
        /// <param name="image">Image rect object to use</param>
        public static Point MoveToImage(Rect image)
        {
            MouseHandler handler = new MouseHandler();
            Point point = Center(image);
            Move(point);

            return Position();
        }

        /// <summary>
        /// Look for and image at the given path and move the cursor at the center of the image
        /// </summary>
        /// <param name="path">Path of the image</param>
        /// <param name="region">Region of the screen to look for the image, for more details look for ScreenRegions class</param>
        public static Point MoveToImage(string path, Rect region = null)
        {
            Rect image = Image.Find(path, region: region);
            MouseHandler handler = new MouseHandler();

            if (image == null)
                return new Point(-1, -1);

            Point point = Center(image);
            Move(point);

            return Position();
        }

        /// <summary>Performs an scroll an amount of ticks from its given position</summary>
        /// <param name="amount">Amount of ticks to scroll the wheel of the mouse</param>
        public static void Scroll(int amount = -500)
        {
            MouseHandler handler = new MouseHandler();
            handler.Wheel(amount);
        }

        ///<summary>Validates if point is a valid point in the screen</summary>
        private static Point FixPoint(Point point)
        {
            if (point.X < 0)
            {
                point.X = 0;
            }
            if (point.Y < 0)
            {
                point.Y = 0;
            }
            return point;
        }

        ///<summary>Gets the center of a rectangle</summary>
        ///<param name="rectangle">rectangle to get the center</param>
        private static Point Center(Rect rectangle)
        {
            Rectangle rect = Image.RectToRectangle(rectangle);
            return new Point(rect.Left + rect.Width / 2, rect.Top + rect.Height / 2);
        }

    }

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

    ///<summary>A class to manipulate image bounds, and alternative to rect struct</summary>
    public class Rect
    {

        public int X { get; set; }
        public int Y { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }

        public Rect(int x, int y, int width, int height)
        {
            this.X = x;
            this.Y = y;
            this.Width = width;
            this.Height = height;
        }

        public override string ToString()
        {
            return string.Format("({0},{1},{2},{3})", X, Y, Width, Height);
        }

    }

    ///<summary>A range object to specified some constraints in images</summary>
    public class Range
    {
        public int Start { get; set; }
        public int End { get; set; }

        public Range(int start, int end)
        {
            Start = start;
            End = end;
        }
    }

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

    ///<summary>A class that provide some methods to manipulate excel files</summary>
    public class Excel
    {
        private XLWorkbook Workbook;
        public int LimitPerSheet { get; set; }
        ///<summary>Set the default limit for rows in an excel sheet</summary>
        public static int DefaultLimit = 200000;
        ///<summary>Get or set the default name for an excel sheet</summary>
        private static string DefaultSheetName = "Sheet1";
        ///<summary>Get or set the default name for an excel file</summary>
        public static string DefaultFileName = "Book.xlsx";
        ///<summary>Get the rows limit for and excel sheet</summary>
        public static int MaxRowLimit { get { return 1048576; } }
        ///<summary>Get all the sheet names in an excel file</summary>
        public ICollection<string> SheetNames { get; private set; }
        ///<summary>Return all the valid object converted types to use in object to row method you can add classes here and override each Tostring method</summary>
        public static IList<Type> ValidConvertedValueTypes = new List<Type>() { typeof(string), typeof(DateTime), typeof(TimeSpan) };
        ///<summary>Return all the supported excell files for this class</summary>
        public static ICollection<string> SupportedExtensions { get { return new string[4] { ".xlsx", ".xslm", ".xltx", ".xltm" }; } }

        ///<summary>Creates and Empty excel file</summary>
        public Excel()
        {
            Workbook = new XLWorkbook(XLEventTracking.Disabled);
            Workbook.Worksheets.Add(DefaultSheetName);
            SheetNames = Workbook.Worksheets.Count > 0 ? Workbook.Worksheets.Select(w => w.Name).ToList() : null;
            LimitPerSheet = DefaultLimit;
            IsValidLimitPerSheet();
        }

        ///<summary>Creates and excel file with a given sheetName and a maximun rows per sheet</summary>
        /// <param name="sheetName">Name of the sheet to assing when the file is created, default Sheet1</param>
        /// <param name="maxRowPerSheet">Amount of rows allowed per file default 200,000 max 1,048,576</param>
        public Excel(string sheetName, int maxRowPerSheet = 200000)
        {
            Workbook = new XLWorkbook(XLEventTracking.Disabled);
            IXLWorksheet newSheet = sheetName == null ? Workbook.Worksheets.Add(DefaultSheetName) : Workbook.Worksheets.Add(sheetName);
            SheetNames = Workbook.Worksheets.Count > 0 ? Workbook.Worksheets.Select(w => w.Name).ToList() : null;
            LimitPerSheet = maxRowPerSheet == 0 ? DefaultLimit : maxRowPerSheet;
            IsValidLimitPerSheet();
        }

        ///<summary>Loads and Excel file</summary>
        /// <param name="path">Path of the file</param>
        /// <param name="maxRowPerSheet">Amount of rows allowed per file default 200,000 max 1,048,576</param>
        public static Excel Load(string path, int maxRowPerSheet = 200000)
        {
            Excel newEx = new Excel();
            newEx.Workbook = new XLWorkbook(path, XLEventTracking.Disabled);
            newEx.SheetNames = newEx.Workbook.Worksheets.Count > 0 ? newEx.Workbook.Worksheets.Select(w => w.Name).ToList() : new List<string>();
            newEx.LimitPerSheet = maxRowPerSheet == 0 ? DefaultLimit : maxRowPerSheet;
            newEx.IsValidLimitPerSheet();
            return newEx;
        }

        ///<summary>Adds a new sheet to the file</summary>
        /// <param name="sheetName">Name of the sheet to add</param>
        public bool AddSheet(string sheetName)
        {
            if (!SheetNames.Contains(sheetName))
            {
                Workbook.Worksheets.Add(sheetName);
                SheetNames.Add(sheetName);
                return true;
            }
            else
                return false;
        }

        ///<summary>Adds a new sheet to the file</summary>
        /// <param name="sheetName">Name of the sheet to add</param>
        /// <param name="table">Data to add to the new sheet</param>
        public bool AddSheet(string sheetName, System.Data.DataTable table)
        {
            if (!SheetNames.Contains(sheetName))
            {
                table.TableName = sheetName;
                Workbook.Worksheets.Add(table);
                SheetNames.Add(sheetName);
                return true;
            }
            else
                return false;
        }

        ///<summary>Removes a sheet of the file</summary>
        /// <param name="sheetName">Name of the sheet to remove, if not exist return false</param>
        public bool RemoveSheet(string sheetName)
        {
            if (SheetNames.Contains(sheetName))
            {
                Workbook.Worksheet(sheetName).Delete();
                SheetNames.Remove(sheetName);
                return true;
            }
            else
                return false;
        }

        ///<summary>Change headers of the file, if the file is new add headers to the first row</summary>
        /// <param name="customHeaders">Headers to add</param>
        /// <param name="sheetName">Name of the sheet to use to add new headers</param>
        public void SetHeaders(IEnumerable<string> customHeaders, string sheetName = null)
        {
            if (customHeaders != null)
            {
                IXLWorksheet workSheet = !string.IsNullOrWhiteSpace(sheetName) ? Workbook.Worksheet(sheetName) : Workbook.Worksheet(1);
                IXLRow row = workSheet.FirstRow();
                int rowIndex = 1;
                foreach (string header in customHeaders)
                {
                    row.Cell(rowIndex).SetValue<string>(header);
                    rowIndex += 1;
                }
            }
        }

        ///<summary>Change headers of the file, if the file is new add headers to the first row</summary>
        /// <param name="Workbook">The current workbook to use to set the header</param>
        /// <param name="customHeaders">Headers to add</param>
        /// <param name="sheetName">Name of the sheet to use to add new headers</param>
        private static void SetHeaders(IXLWorkbook Workbook, IEnumerable<string> customHeaders, string sheetName = null)
        {
            if (customHeaders != null)
            {
                IXLWorksheet workSheet = !string.IsNullOrWhiteSpace(sheetName) ? Workbook.Worksheet(sheetName) : Workbook.Worksheet(1);
                IXLRow row = workSheet.FirstRow();
                int rowIndex = 1;
                foreach (string header in customHeaders)
                {
                    row.Cell(rowIndex).SetValue<string>(header);
                    rowIndex += 1;
                }
            }
        }

        private static void SetHeaders(IEnumerable<string> headers, IXLWorksheet workSheet)
        {
            if (headers != null)
            {
                int rowIndex = 1;
                IXLRow row = workSheet.FirstRow();
                foreach (string header in headers)
                {
                    row.Cell(rowIndex).SetValue<string>(header);
                    rowIndex += 1;
                }
            }
        }

        /// <summary>
        /// Return the current headers of the file
        /// </summary>
        /// <param name="sheetName">Name of the sheet to look for headers</param>
        public IEnumerable<string> GetHeaders(string sheetName = null)
        {
            IXLWorksheet workSheet = !string.IsNullOrWhiteSpace(sheetName) ? Workbook.Worksheet(sheetName) : Workbook.Worksheet(1);
            return workSheet.FirstRow().CellsUsed(XLCellsUsedOptions.Contents).Select(c => c.GetString()).ToList();
        }

        private IEnumerable<string> GetHeaders(IXLWorksheet workSheet)
        {
            return workSheet.FirstRow().CellsUsed(XLCellsUsedOptions.Contents).Select(c => c.GetString()).ToList();
        }

        ///<summary>Get all data of the current excel file sheet as a list of dictionaries</summary>
        /// <param name="sheetName">Name of the sheet to get the data, if null it will use the first sheet</param>
        /// <param name="customHeaders">Headers to use as keys to the dictionary that represends a row</param>
        public IEnumerable<IDictionary<string, object>> Extract(string sheetName = null, IEnumerable<string> customHeaders = null)
        {
            IList<IDictionary<string, object>> dictList = new List<IDictionary<string, object>>();
            IXLWorksheet workSheet = string.IsNullOrWhiteSpace(sheetName) ? Workbook.Worksheet(1) : Workbook.Worksheet(sheetName);
            IList<string> headers = (customHeaders == null ? GetHeaders(workSheet) : customHeaders).ToList();

            foreach (IXLRow row in workSheet.RowsUsed(XLCellsUsedOptions.Contents).Skip(customHeaders == null ? 1 : 0))
            {
                IDictionary<string, object> dict = new Dictionary<string, object>();
                for (int i = 0; i < headers.Count; i++)
                {
                    dict.Add(headers[i], row.Cell(i + 1).Value);
                }
                dictList.Add(dict);
            }
            return dictList;
        }

        ///<summary>Get all data of the current excel file sheet as a list of dictionaries starting and ending from the specified params</summary>
        ///<param name="startRow">Row to start to take the data</param>
        /// <param name="endRow">Row to stop to take the data</param>
        /// <param name="sheetName">Name of the sheet to get the data, if null it will use the first sheet</param>
        /// <param name="customHeaders">Headers to use as keys to the dictionary that represends a row</param>
        public IEnumerable<IDictionary<string, object>> Extract(int startRow, int endRow, string sheetName = null, IEnumerable<string> customHeaders = null)
        {
            IList<IDictionary<string, object>> dictList = new List<IDictionary<string, object>>();
            IXLWorksheet workSheet = string.IsNullOrWhiteSpace(sheetName) ? Workbook.Worksheet(1) : Workbook.Worksheet(sheetName);
            IList<string> headers = (customHeaders == null ? GetHeaders(workSheet) : customHeaders).ToList();

            foreach (IXLRow row in workSheet.RowsUsed(XLCellsUsedOptions.Contents).Skip(customHeaders == null ? 1 : 0).Skip(startRow).Take(endRow))
            {
                IDictionary<string, object> dict = new Dictionary<string, object>();
                for (int i = 0; i < headers.Count; i++)
                {
                    dict.Add(headers[i], row.Cell(i + 1).Value);
                }
                dictList.Add(dict);
            }
            return dictList;
        }

        ///<summary>Get all data of the current excel file for all sheets as list of list of dictionaries where each list represents a sheet</summary>
        /// <param name="customHeaders">Headers to use as keys to the dictionary that represends a row</param>
        public IEnumerable<IEnumerable<IDictionary<string, object>>> ExtractAll(IEnumerable<string> customHeaders = null)
        {
            ICollection<IEnumerable<IDictionary<string, object>>> listOfListDict = new List<IEnumerable<IDictionary<string, object>>>();
            foreach (string name in SheetNames)
            {
                listOfListDict.Add(Extract(name, customHeaders));
            }
            return listOfListDict;
        }

        /// <summary>Get all data of the current excel file  as a datatable</summary>
        /// <param name="sheetName">Name of the sheet to get the data, if null it will use the first sheet</param>
        /// <param name="customHeaders">Headers to use as keys to the dictionary that represends a row</param>
        public System.Data.DataTable ExtractToDatatable(string sheetName = null, IEnumerable<string> customHeaders = null)
        {
            System.Data.DataTable table = new System.Data.DataTable();
            IXLWorksheet workSheet = string.IsNullOrWhiteSpace(sheetName) ? Workbook.Worksheet(1) : Workbook.Worksheet(sheetName);
            IEnumerable<string> headers = customHeaders == null ? GetHeaders(workSheet) : customHeaders;

            table.TableName = workSheet.Name;

            foreach (string header in headers)
                table.Columns.Add(header);

            foreach (IXLRow row in workSheet.RowsUsed().Skip(customHeaders == null ? 1 : 0))
            {
                System.Data.DataRow dataRow = table.NewRow();
                dataRow.ItemArray = row.CellsUsed(XLCellsUsedOptions.Contents).Select(c => c.Value).ToArray();
                table.Rows.Add(dataRow);
            }
            return table;
        }

        ///<summary>Get all data of the current excel file sheet as a datatable starting and ending from the specified params</summary>
        ///<param name="startRow">Row to start to take the data</param>
        /// <param name="endRow">Row to stop to take the data</param>
        /// <param name="sheetName">Name of the sheet to get the data, if null it will use the first sheet</param>
        /// <param name="customHeaders">Headers to use as keys to the dictionary that represends a row</param> 
        public System.Data.DataTable ExtractToDatatable(int startRow, int endRow, string sheetName = null, IEnumerable<string> customHeaders = null)
        {
            System.Data.DataTable table = new System.Data.DataTable();
            IXLWorksheet workSheet = string.IsNullOrWhiteSpace(sheetName) ? Workbook.Worksheet(1) : Workbook.Worksheet(sheetName);
            IEnumerable<string> headers = customHeaders == null ? GetHeaders(workSheet) : customHeaders;

            table.TableName = workSheet.Name;

            foreach (string header in headers)
                table.Columns.Add(header);

            foreach (IXLRow row in workSheet.RowsUsed().Skip(customHeaders == null ? 1 : 0).Skip(startRow).Take(endRow))
            {
                System.Data.DataRow dataRow = table.NewRow();
                dataRow.ItemArray = row.CellsUsed(XLCellsUsedOptions.Contents).Select(c => c.Value).ToArray();
                table.Rows.Add(dataRow);
            }
            return table;
        }

        ///<summary>Get all data of the current excel file for all sheets as list of datatable where each datatable represents a sheet</summary>
        /// <param name="customHeaders">Headers to use as keys to the dictionary that represends a row</param>
        public IEnumerable<System.Data.DataTable> ExtractAllToDatatable(IEnumerable<string> customHeaders = null)
        {
            ICollection<System.Data.DataTable> listOfDatatable = new List<System.Data.DataTable>();
            foreach (string name in SheetNames)
            {
                listOfDatatable.Add(ExtractToDatatable(name, customHeaders));
            }
            return listOfDatatable;
        }

        ///<summary>Get all data of the current excel file sheet and cast it to an specified object if possible</summary>
        /// <param name="sheetName">Name of the sheet to get the data, if null it will use the first sheet</param>
        /// <param name="customHeaders">Headers to use to reflect the objets variable names of your class</param>
        public IEnumerable<T> ExtractToObject<T>(string sheetName = null, IEnumerable<string> customHeaders = null) where T : new()
        {
            IList<T> objectList = new List<T>();
            IXLWorksheet workSheet = string.IsNullOrWhiteSpace(sheetName) ? Workbook.Worksheet(1) : Workbook.Worksheet(sheetName);
            IList<string> headers = (customHeaders == null ? GetHeaders(workSheet) : customHeaders).ToList();

            foreach (IXLRow row in workSheet.RowsUsed().Skip(customHeaders == null ? 1 : 0))
            {
                objectList.Add(GetObjectFromRow<T>(headers, row));
            }
            return objectList;
        }

        ///<summary>Get all data of the current excel file for sheet as a list of objets starting and ending from the specified params</summary>
        ///<param name="startRow">Row to start to take the data</param>
        /// <param name="endRow">Row to stop to take the data</param>
        /// <param name="sheetName">Name of the sheet to get the data, if null it will use the first sheet</param>
        /// <param name="customHeaders">Headers to use as keys to the dictionary that represends a row</param>
        public IEnumerable<T> ExtractToObject<T>(int startRow, int endRow, string sheetName = null, IEnumerable<string> customHeaders = null) where T : new()
        {
            IList<T> objectList = new List<T>();
            IXLWorksheet workSheet = string.IsNullOrWhiteSpace(sheetName) ? Workbook.Worksheet(1) : Workbook.Worksheet(sheetName);
            IList<string> headers = (customHeaders == null ? GetHeaders(workSheet) : customHeaders).ToList();

            foreach (IXLRow row in workSheet.RowsUsed().Skip(customHeaders == null ? 1 : 0).Skip(startRow).Take(endRow))
            {
                objectList.Add(GetObjectFromRow<T>(headers, row));
            }
            return objectList;
        }

        ///<summary>Get all data of the current excel file for all sheets as a list of list of objects, where each list represents a sheet</summary>
        /// <param name="customHeaders">Headers to use as keys to the dictionary that represends a row</param>
        public IEnumerable<IEnumerable<T>> ExtractAllToObject<T>(IEnumerable<string> customHeaders = null) where T : new()
        {
            ICollection<IEnumerable<T>> listOfList = new List<IEnumerable<T>>();
            foreach (string name in SheetNames)
            {
                listOfList.Add(ExtractToObject<T>(name, customHeaders));
            }
            return listOfList;
        }

        ///<summary>Get a column of the specified file sheet</summary>
        /// <param name="column">Name of the column title in the excel file sheet</param>
        /// <param name="sheetName">Name of the sheet to get the data, if null it will use the first sheet</param>
        public IEnumerable<T> GetColumn<T>(string column, string sheetName = null) where T : new()
        {
            IXLWorksheet workSheet = string.IsNullOrWhiteSpace(sheetName) ? Workbook.Worksheet(1) : Workbook.Worksheet(sheetName);
            List<string> headers = GetHeaders(workSheet).ToList();
            int index = headers.IndexOf(column);

            if (index < 0)
                throw new Exception(string.Format("The column {0} is not valid", column));

            return workSheet.Column(index).CellsUsed(XLCellsUsedOptions.Contents).Skip(1).Select(c => c.GetValue<T>()).ToList();
        }

        ///<summary>Get a column of the specified file sheet</summary>
        /// <param name="column">Index of the column starting from 1 in the excel file sheet</param>
        /// <param name="sheetName">Name of the sheet to get the data, if null it will use the first sheet</param>
        public IEnumerable<T> GetColumn<T>(int column, string sheetName = null) where T : new()
        {
            ICollection<string> values = new List<string>();
            IXLWorksheet workSheet = string.IsNullOrWhiteSpace(sheetName) ? Workbook.Worksheet(1) : Workbook.Worksheet(sheetName);
            IXLRow headers = workSheet.FirstRow();

            if (column < 1 || column > workSheet.ColumnsUsed().Count())
                throw new Exception(string.Format("The column index {0} is not valid", column));

            return workSheet.Column(column).CellsUsed(XLCellsUsedOptions.Contents).Skip(1).Select(c => c.GetValue<T>()).ToList();
        }

        ///<summary>Creates a new excel file</summary>
        /// <param name="path">Filepath of the new excel file</param>
        /// <param name="rows">Data to add to the new excel file as list of list where each list repressent a row</param>
        /// <param name="headers">Titles of the excel file, if null it will start adding rows from the first file</param>
        /// <param name="sheetName">Name of the sheet to get the data, if null it will use the first sheet</param>
        public static void Create(string path, IEnumerable<object[]> rows, IEnumerable<string> headers = null, string sheetName = null)
        {
            string folder = Path.GetDirectoryName(path);
            ValidateFolder(folder);

            int iterations = 0;
            int amountOfItems = rows.Count();
            int amountOfItemsToSkip = DefaultLimit * iterations;
            XLWorkbook newWorkBook = new XLWorkbook();
            IXLWorksheet workSheet = sheetName == null ? newWorkBook.Worksheets.Add(DefaultSheetName) : newWorkBook.Worksheets.Add(sheetName);
            int currentDefaultLimit = DefaultLimit == MaxRowLimit ? DefaultLimit - 1 : DefaultLimit;

            SetHeaders(headers, workSheet);

            while (amountOfItemsToSkip < amountOfItems)
            {
                workSheet.Cell(2, 1).InsertData(rows.Skip(amountOfItemsToSkip).Take(currentDefaultLimit));
                iterations++;
                amountOfItemsToSkip = currentDefaultLimit * iterations;

                if (amountOfItemsToSkip >= amountOfItems)
                    break;

                workSheet = PrepareNewSheet(newWorkBook, workSheet);
            }

            newWorkBook.SaveAs(path);
        }

        ///<summary>Creates a new excel file</summary>
        /// <param name="path">Filepath of the new excel file</param>
        /// <param name="rows">Data to add to the new excel file as list of object with each object represents a row</param>
        /// <param name="headers">Titles of the excel file, if null it will start adding rows from the first file</param>
        /// <param name="sheetName">Name of the sheet to get the data, if null it will use the first sheet</param>
        public static void CreateFromObject<T>(string path, IEnumerable<T> rows, string sheetName = null) where T : class
        {
            string folder = Path.GetDirectoryName(path);
            ValidateFolder(folder);

            int iterations = 0;
            int amountOfItems = rows.Count();
            int amountOfItemsToSkip = DefaultLimit * iterations;
            XLWorkbook newWorkBook = new XLWorkbook();

            IXLWorksheet workSheet = sheetName == null ? newWorkBook.Worksheets.Add(DefaultSheetName) : newWorkBook.Worksheets.Add(sheetName);
            int currentDefaultLimit = DefaultLimit == MaxRowLimit ? DefaultLimit - 1 : DefaultLimit;

            while (amountOfItemsToSkip < amountOfItems)
            {
                workSheet.Cell(1, 1).InsertData(rows.Skip(amountOfItemsToSkip).Take(currentDefaultLimit));

                iterations++;
                amountOfItemsToSkip = currentDefaultLimit * iterations;

                if (amountOfItemsToSkip >= amountOfItems)
                    break;

                workSheet = PrepareNewSheet(newWorkBook, workSheet);
            }
            newWorkBook.SaveAs(path);
        }

        ///<summary>Creates a new excel file</summary>
        /// <param name="path">Filepath of the new excel file</param>
        /// <param name="rows">Data to add to the new excel file as list of dictionaries where each dictionary represents a row</param>
        /// <param name="headers">Titles of the excel file, if null it will start adding rows from the first file</param>
        /// <param name="sheetName">Name of the sheet to get the data, if null it will use the first sheet</param>
        public static void Create(string path, IEnumerable<IDictionary<string, object>> rows, string sheetName = null)
        {
            string folder = Path.GetDirectoryName(path);
            ValidateFolder(folder);

            int iterations = 0;
            int amountOfItems = rows.Count();
            int amountOfItemsToSkip = DefaultLimit * iterations;
            XLWorkbook newWorkBook = new XLWorkbook();
            IXLWorksheet workSheet = sheetName == null ? newWorkBook.Worksheets.Add(DefaultSheetName) : newWorkBook.Worksheets.Add(sheetName);
            int currentDefaultLimit = DefaultLimit == MaxRowLimit ? DefaultLimit - 1 : DefaultLimit;
            IDictionary<string, object> firstRow = rows.FirstOrDefault();

            if (firstRow == null)
                throw new Exception("There must be at least 1 row in rows data");

            AddHeaders(workSheet, firstRow.Keys);

            while (amountOfItemsToSkip < amountOfItems)
            {
                workSheet.Cell(2, 1).InsertData(rows.Skip(amountOfItemsToSkip).Take(currentDefaultLimit).Select(r => r.Values.ToArray()));
                iterations++;
                amountOfItemsToSkip = currentDefaultLimit * iterations;

                if (amountOfItemsToSkip >= amountOfItems)
                    break;

                workSheet = PrepareNewSheet(newWorkBook, workSheet);
            }

            newWorkBook.SaveAs(path);
        }

        ///<summary>Creates a new excel file</summary>
        /// <param name="path">Filepath of the new excel file</param>
        /// <param name="table">Data to add to the new excel file as a datatable</param>
        /// <param name="headers">Titles of the excel file, if null it will start adding rows from the first file</param>
        /// <param name="sheetName">Name of the sheet to get the data, if null it will use the first sheet</param>
        public static void Create(string path, System.Data.DataTable table, string sheetName = null)
        {
            string folder = Path.GetDirectoryName(path);
            ValidateFolder(folder);

            XLWorkbook newWorkBook = new XLWorkbook();
            IXLWorksheet workSheet = sheetName == null ? newWorkBook.Worksheets.Add(DefaultSheetName) : newWorkBook.Worksheets.Add(sheetName);
            IXLTable xltable = workSheet.Cell(1, 1).InsertTable(table);
            RemoveTableStyles(xltable);

            newWorkBook.SaveAs(path);
        }

        ///<summary>Add data to an excel file</summary>
        /// <param name="row">Data to add to the new excel file as list</param>
        /// <param name="sheetName">Name of the sheet to get the data, if null it will use the first sheet</param>
        public void Append(IEnumerable<object> row, string sheetName = null)
        {
            IXLRow rowSheet = null;
            IXLWorksheet workSheet = sheetName == null ? Workbook.Worksheet(1) : Workbook.Worksheet(sheetName);
            IXLRow lastRow = workSheet.LastRowUsed(XLCellsUsedOptions.Contents);
            int rowIndex = lastRow == null ? 1 : lastRow.RowNumber() + 1;
            int cellIndex = 1;

            if (rowIndex > LimitPerSheet)
            {
                workSheet = PrepareNewSheet(Workbook, workSheet);
                rowIndex = 2;
            }

            rowSheet = workSheet.Row(rowIndex);
            foreach (object value in row)
            {
                rowSheet.Cell(cellIndex).Value = value;
                cellIndex += 1;
            }

        }

        ///<summary>Add data to an excel file</summary>
        /// <param name="rows">Data to add to the new excel file as list of list where each list represents a row</param>
        /// <param name="sheetName">Name of the sheet to get the data, if null it will use the first sheet</param>
        public void Append(IEnumerable<object[]> rows, string sheetName = null)
        {
            int iterations = 0;
            int amountOfItems = rows.Count();
            int amountOfItemsToSkip = LimitPerSheet * iterations;
            IXLWorksheet workSheet = sheetName == null ? Workbook.Worksheet(1) : Workbook.Worksheet(sheetName);
            IXLRow lastRow = workSheet.LastRowUsed(XLCellsUsedOptions.Contents);
            int rowIndex = lastRow == null ? 1 : lastRow.RowNumber() + 1;

            while (amountOfItemsToSkip < amountOfItems)
            {
                workSheet.Cell(rowIndex, 1).InsertData(rows.Skip(amountOfItemsToSkip).Take(LimitPerSheet));
                iterations++;
                amountOfItemsToSkip = LimitPerSheet * iterations;

                if (amountOfItemsToSkip >= amountOfItems)
                    break;

                workSheet = PrepareNewSheet(Workbook, workSheet);
            }

        }

        ///<summary>Add data to an excel file</summary>
        /// <param name="row">Data to add to the new excel file as and object</param>
        /// <param name="sheetName">Name of the sheet to get the data, if null it will use the first sheet</param>
        public void AppendFromObject<T>(T row, string sheetName = null) where T : class
        {
            IXLWorksheet workSheet = sheetName == null ? Workbook.Worksheet(1) : Workbook.Worksheet(sheetName);
            IXLRow lastRow = workSheet.LastRowUsed(XLCellsUsedOptions.Contents);
            int rowIndex = lastRow == null ? 1 : lastRow.RowNumber() + 1;
            int cellIndex = 1;
            IXLRow rowSheet = null;

            if (rowIndex > LimitPerSheet)
            {
                workSheet = PrepareNewSheet(Workbook, workSheet);
                rowIndex = 2;
            }

            rowSheet = workSheet.Row(rowIndex);
            foreach (PropertyInfo prop in row.GetType().GetProperties())
            {
                if (IsValidPropertyType(prop.PropertyType))
                {
                    rowSheet.Cell(cellIndex).Value = prop.GetValue(row, null);
                    cellIndex++;
                }
                rowIndex++;
            }
        }

        ///<summary>Add data to an excel file</summary>
        /// <param name="rows">Data to add to the new excel file as a list of objects where each object represents a row</param>
        /// <param name="sheetName">Name of the sheet to get the data, if null it will use the first sheet</param>
        public void AppendFromObject<T>(IEnumerable<T> rows, string sheetName = null) where T : class
        {
            int iterations = 0;
            int amountOfItems = rows.Count();
            int amountOfItemsToSkip = LimitPerSheet * iterations;
            IXLWorksheet workSheet = sheetName == null ? Workbook.Worksheets.First() : Workbook.Worksheets.Worksheet(sheetName);
            int rowIndex = workSheet.LastRowUsed().RowNumber() + 1;

            while (amountOfItemsToSkip < amountOfItems)
            {
                workSheet.Cell(rowIndex, 1).InsertData(rows.Skip(amountOfItemsToSkip).Take(LimitPerSheet));

                iterations++;
                amountOfItemsToSkip = LimitPerSheet * iterations;

                if (amountOfItemsToSkip >= amountOfItems)
                    break;

                workSheet = PrepareNewSheet(Workbook, workSheet);
            }
        }

        ///<summary>Add data to an excel file</summary>
        /// <param name="row">Data to add to the new excel file as a dictionary</param>
        /// <param name="sheetName">Name of the sheet to get the data, if null it will use the first sheet</param>
        public void Append(IDictionary<string, object> row, string sheetName = null)
        {
            IXLRow rowSheet = null;
            IXLWorksheet workSheet = sheetName == null ? Workbook.Worksheets.First() : Workbook.Worksheets.Worksheet(sheetName);

            IXLRow firstRow = workSheet.FirstRowUsed(XLCellsUsedOptions.Contents);
            IXLRow lastRow = workSheet.LastRowUsed(XLCellsUsedOptions.Contents);

            int rowIndex = lastRow == null ? 1 : lastRow.RowNumber() + 1;
            int cellIndex = 1;

            if (firstRow == null)
            {
                SetHeaders(row.Keys);
                rowIndex++;
            }

            if (rowIndex > LimitPerSheet)
            {
                workSheet = PrepareNewSheet(Workbook, workSheet);
                rowIndex = 2;
            }

            rowSheet = workSheet.Row(rowIndex);
            foreach (object value in row.Values)
            {
                rowSheet.Cell(cellIndex).Value = value;
                cellIndex++;
            }

        }

        ///<summary>Add data to an excel file</summary>
        /// <param name="rows">Data to add to the new excel file as a list of dictionaries where each dictionary represents a row</param>
        /// <param name="sheetName">Name of the sheet to get the data, if null it will use the first sheet</param>
        public void Append(IEnumerable<IDictionary<string, object>> rows, string sheetName = null)
        {
            int iterations = 0;
            int amountOfItems = rows.Count();
            int amountOfItemsToSkip = LimitPerSheet * iterations;
            IXLWorksheet workSheet = sheetName == null ? Workbook.Worksheets.First() : Workbook.Worksheets.Worksheet(sheetName);

            IXLRow firstRow = workSheet.FirstRowUsed(XLCellsUsedOptions.Contents);
            IXLRow lastRow = workSheet.LastRowUsed(XLCellsUsedOptions.Contents);

            int rowIndex = lastRow == null ? 1 : lastRow.RowNumber() + 1;

            if (firstRow == null && amountOfItems > 0)
            {
                SetHeaders(rows.First().Keys);
                rowIndex++;
            }

            while (amountOfItemsToSkip < amountOfItems)
            {
                workSheet.Cell(rowIndex, 1).InsertData(rows.Skip(amountOfItemsToSkip).Take(LimitPerSheet).Select(r => r.Values.ToArray()));

                iterations++;
                amountOfItemsToSkip = LimitPerSheet * iterations;

                if (amountOfItemsToSkip >= amountOfItems)
                    break;

                workSheet = PrepareNewSheet(Workbook, workSheet);
            }
        }

        ///<summary>Validates a header in a</summary>
        /// <param name="header">Header to validate in the excel file</param>
        /// <param name="sheetName">Name of the sheet to get the data, if null it will use the first sheet</param>
        /// <param name="exactMatch">If is true it will validate the string without case sensitive</param>
        public bool HasHeader(string header, string sheetName = null, bool exactMatch = true)
        {
            IXLWorksheet workSheet = string.IsNullOrWhiteSpace(sheetName) ? Workbook.Worksheet(1) : Workbook.Worksheet(sheetName);
            IXLRow firstRow = workSheet.FirstRow();
            IEnumerable<string> cells = null;
            string headerCopy = exactMatch ? header : header.ToLower();

            if (firstRow == null)
                return false;

            cells = firstRow.CellsUsed(XLCellsUsedOptions.Contents).Select(c => c.GetString());
            if (cells.Count() == 0)
                return false;

            return exactMatch ? cells.Any(c => c == headerCopy) : cells.Any(c => c.ToLower() == headerCopy);
        }

        ///<summary>Add data to an excel file</summary>
        /// <param name="header">List of headers to validate in the excel file</param>
        /// <param name="sheetName">Name of the sheet to get the data, if null it will use the first sheet</param>
        /// <param name="exactMatch">If is true it will validate the string without case sensitive</param>
        public bool HasHeader(IEnumerable<string> headers, string sheetName = null, bool exactMatch = true)
        {
            IXLWorksheet workSheet = string.IsNullOrWhiteSpace(sheetName) ? Workbook.Worksheet(1) : Workbook.Worksheet(sheetName);
            IXLRow firstRow = workSheet.FirstRow();
            IEnumerable<string> cells = null;

            if (firstRow == null)
                return false;

            cells = exactMatch ? firstRow.CellsUsed(XLCellsUsedOptions.Contents).Select(c => c.GetString()) : firstRow.CellsUsed(XLCellsUsedOptions.Contents).Select(c => c.GetString().ToLower());
            if (cells.Count() == 0)
                return false;

            foreach (string header in (exactMatch ? headers : headers.Select(h => h.ToLower())))
            {
                if (!cells.Contains(header))
                    return false;
            }

            return true;
        }

        ///<summary>Add data to an excel file</summary>
        /// <param name="obj">Object to use to compare if sheet contains the object variables</param>
        /// <param name="sheetName">Name of the sheet to get the data, if null it will use the first sheet</param>
        /// <param name="exactMatch">If is true it will validate the string without case sensitive</param>
        public bool HasHeaderFromObject<T>(T obj, string sheetName = null, bool exactMatch = true) where T : class
        {
            IXLWorksheet workSheet = string.IsNullOrWhiteSpace(sheetName) ? Workbook.Worksheet(1) : Workbook.Worksheet(sheetName);
            IXLRow firstRow = workSheet.FirstRow();
            IEnumerable<string> cells = null;
            IEnumerable<string> headers = obj.GetType().GetProperties().Where(p => IsValidPropertyType(p.PropertyType)).Select(p => p.Name);

            if (firstRow == null)
                return false;

            cells = exactMatch ? firstRow.CellsUsed(XLCellsUsedOptions.Contents).Select(c => c.GetString()) : firstRow.CellsUsed().Select(c => c.GetString().ToLower());
            if (cells.Count() == 0)
                return false;

            foreach (string header in (exactMatch ? headers : headers.Select(h => h.ToLower())))
            {
                if (!cells.Contains(header))
                    return false;
            }

            return true;
        }

        ///<summary>Clear all rows of a sheet except for headers</summary>
        /// <param name="sheetName">Name of the sheet to get the data, if null it will use the first sheet</param>
        public void Reset(string sheetName = null)
        {
            IXLWorksheet workSheet = string.IsNullOrWhiteSpace(sheetName) ? Workbook.Worksheet(1) : Workbook.Worksheet(sheetName);
            IEnumerable<string> headers = GetHeaders(workSheet);

            workSheet.Clear(XLClearOptions.AllContents);
            SetHeaders(headers, workSheet);
        }

        ///<summary>Clear all rows of a sheet incliding header</summary>
        /// <param name="sheetName">Name of the sheet to get the data, if null it will use the first sheet</param>
        public void Clear(string sheetName = null)
        {
            IXLWorksheet workSheet = string.IsNullOrWhiteSpace(sheetName) ? Workbook.Worksheet(1) : Workbook.Worksheet(sheetName);
            workSheet.Clear(XLClearOptions.AllContents);
        }

        ///<summary>Writes data to an excel file</summary>
        /// <param name="path">Filepath of the new excel file</param>
        /// <param name="rows">Data to add to the new excel file as a list of objects where each object represents a row</param>
        /// <param name="sheetName">Name of the sheet to get the data, if null it will use the first sheet</param>
        public static void Write(string path, IEnumerable<IEnumerable<string>> rows, string sheetName = null)
        {
            int iterations = 0;
            int amountOfItems = rows.Count();
            int amountOfItemsToSkip = DefaultLimit * iterations;
            XLWorkbook currentWorkBook = new XLWorkbook(path);

            IXLWorksheet workSheet = sheetName == null ? currentWorkBook.Worksheet(1) : currentWorkBook.Worksheet(sheetName);
            IXLRow lastRow = workSheet.LastRowUsed(XLCellsUsedOptions.Contents);
            int rowIndex = (lastRow == null) ? 1 : lastRow.RowNumber() + 1;

            while (amountOfItemsToSkip < amountOfItems)
            {
                workSheet.Cell(rowIndex, 1).InsertData(rows.Skip(amountOfItemsToSkip).Take(DefaultLimit));
                iterations++;
                amountOfItemsToSkip = DefaultLimit * iterations;

                if (amountOfItemsToSkip >= amountOfItems)
                    break;

                workSheet = PrepareNewSheet(currentWorkBook, workSheet);
            }

            Save(currentWorkBook, path);
        }

        /// <summary>Writes data to an excel file</summary>
        /// <param name="path">Filepath of the new excel file</param>
        /// <param name="rows">Data to add to the new excel file as a list of objects where each object represents a row</param>
        /// <param name="sheetName">Name of the sheet to get the data, if null it will use the first sheet</param>
        public static void WriteFromObject<T>(string path, IEnumerable<T> rows, string sheetName = null) where T : class
        {
            ValidateFile(path);

            XLWorkbook currentWorkBook = new XLWorkbook(path);
            int iterations = 0;
            int amountOfItems = rows.Count();
            int amountOfItemsToSkip = DefaultLimit * iterations;
            IXLWorksheet workSheet = sheetName == null ? currentWorkBook.Worksheet(1) : currentWorkBook.Worksheets.Worksheet(sheetName);
            IXLRow lastRow = workSheet.LastRowUsed(XLCellsUsedOptions.Contents);
            int rowIndex = (lastRow == null) ? 1 : lastRow.RowNumber() + 1;

            while (amountOfItemsToSkip < amountOfItems)
            {
                workSheet.Cell(rowIndex, 1).InsertData(rows.Skip(amountOfItemsToSkip).Take(DefaultLimit));

                iterations++;
                amountOfItemsToSkip = DefaultLimit * iterations;

                if (amountOfItemsToSkip >= amountOfItems)
                    break;

                workSheet = PrepareNewSheet(currentWorkBook, workSheet);
            }
            Save(currentWorkBook, path);
        }

        /// <summary>Writes data to an excel file</summary>
        /// <param name="path">Filepath of the new excel file</param>
        /// <param name="rows">Data to add to the new excel file as a list of dictionaries where each dictionary represents a row</param>
        /// <param name="sheetName">Name of the sheet to get the data, if null it will use the first sheet</param>
        public static void Write(string path, IEnumerable<IDictionary<string, string>> rows, string sheetName = null)
        {
            ValidateFile(path);

            XLWorkbook currentWorkBook = new XLWorkbook(path, XLEventTracking.Disabled);
            int iterations = 0;
            int amountOfItems = rows.Count();
            int amountOfItemsToSkip = DefaultLimit * iterations;
            IXLWorksheet workSheet = sheetName == null ? currentWorkBook.Worksheet(1) : currentWorkBook.Worksheets.Worksheet(sheetName);

            IXLRow firstRow = workSheet.FirstRowUsed(XLCellsUsedOptions.Contents);
            IXLRow lastRow = workSheet.LastRowUsed(XLCellsUsedOptions.Contents);

            int rowIndex = lastRow == null ? 1 : lastRow.RowNumber() + 1;

            if (firstRow == null && amountOfItems > 0)
            {
                SetHeaders(currentWorkBook, rows.First().Keys);
                rowIndex++;
            }

            while (amountOfItemsToSkip < amountOfItems)
            {
                workSheet.Cell(rowIndex, 1).InsertData(rows.Skip(amountOfItemsToSkip).Take(DefaultLimit).Select(r => r.Values.ToArray()));

                iterations++;
                amountOfItemsToSkip = DefaultLimit * iterations;

                if (amountOfItemsToSkip >= amountOfItems)
                    break;

                workSheet = PrepareNewSheet(currentWorkBook, workSheet);
            }
            Save(currentWorkBook, path);
        }

        /// <summary>Split an excel file sheet into multiple files by an specified quantity</summary>
        /// <param name="path">Filepath of the new excel file</param>
        /// <param name="newFilesfolderPath">Folder to use to store the generated files</param>
        /// <param name="filesQuantity">Amount of files to create</param>
        /// <param name="sheetName">Name of the sheet to get the data, if null it will use the first sheet</param>
        /// <param name="customHeaders">Titles to use for each file created</param>
        /// <param name="template_name">Name to use for all the files, ej: if Book then files will be Book1, Book2 etc..</param>
        public static void SplitByFiles(string path, string newFilesfolderPath, int filesQuantity, string sheetName = null, IEnumerable<string> customHeaders = null, string template_name = null)
        {
            ValidateFile(path);
            ValidateFolder(newFilesfolderPath);

            XLWorkbook CurrentWorkBook = new XLWorkbook(path, XLEventTracking.Disabled);
            string filename = string.IsNullOrEmpty(template_name) ? Path.GetFileNameWithoutExtension(DefaultFileName) : Path.GetFileNameWithoutExtension(template_name);
            string extension = string.IsNullOrEmpty(template_name) ? Path.GetExtension(DefaultFileName) : Path.GetExtension(template_name);

            if (!IsValidExtension(extension))
                throw new Exception(string.Format("The extension {0} is not valid, valid extensions:{1}", extension, string.Join(",", SupportedExtensions)));

            IXLWorksheet workSheet = sheetName == null ? CurrentWorkBook.Worksheet(1) : CurrentWorkBook.Worksheet(sheetName);
            IEnumerable<IXLRow> rows = workSheet.RowsUsed(XLCellsUsedOptions.Contents, r => r.RowNumber() > 1);
            IList<string> headers = (customHeaders == null ? workSheet.FirstRowUsed(XLCellsUsedOptions.Contents).CellsUsed(XLCellsUsedOptions.Contents).Select(c => c.GetString()) : customHeaders).ToList();

            int fileRowsNumber = workSheet.LastRowUsed(XLCellsUsedOptions.Contents).RowNumber() - 1;
            int remainderRow = fileRowsNumber % filesQuantity;
            int adittionalRowsPerSheet = remainderRow > 0 ? Convert.ToInt32(Math.Ceiling((double)remainderRow / filesQuantity)) : 0;
            int totalRows = (fileRowsNumber / filesQuantity) + adittionalRowsPerSheet;

            int filesIndexName = 1;
            int iterations = 0;

            XLWorkbook newWorbook = new XLWorkbook(XLEventTracking.Disabled);
            IXLWorksheet newWorksheet = newWorbook.Worksheets.Add(DefaultSheetName);
            string fullFileName = Path.Combine(newFilesfolderPath, filename + filesIndexName.ToString() + extension);

            SetHeaders(newWorbook, headers, newWorksheet.Name);

            while (iterations < filesQuantity)
            {
                IEnumerable<IXLRow> currentRows = rows.Skip((totalRows * iterations)).Take(totalRows);
                ICollection<object[]> currentCells = new List<object[]>();

                foreach (IXLRow item in currentRows)
                {
                    currentCells.Add(item.CellsUsed(XLCellsUsedOptions.Contents).Select(c => c.Value).ToArray());
                }

                newWorksheet.Cell(2, 1).InsertData(currentCells);

                Save(newWorbook, fullFileName);
                newWorbook = new XLWorkbook(XLEventTracking.Disabled);
                newWorksheet = newWorbook.Worksheets.Add(DefaultSheetName);
                SetHeaders(newWorbook, headers, newWorksheet.Name);

                filesIndexName++;
                iterations++;

                fullFileName = Path.Combine(newFilesfolderPath, filename + filesIndexName.ToString() + extension);
            }
        }

        /// <summary>Split an excel file sheet into multiple files by an specified amount of rows</summary>
        /// <param name="path">Filepath of the new excel file</param>
        /// <param name="newFilesfolderPath">Folder to use to store the generated files</param>
        /// <param name="rowsQuantity">Amount of rows per file </param>
        /// <param name="sheetName">Name of the sheet to get the data, if null it will use the first sheet</param>
        /// <param name="customHeaders">Titles to use for each file created</param>
        /// <param name="template_name">Name to use for all the files, ej: if Book then files will be Book1, Book2 etc..</param>
        public static void SplitByRows(string path, string newFilesfolderPath, int rowsQuantity, string sheetName = null, IEnumerable<string> customHeaders = null, string template_name = null)
        {
            ValidateFile(path);
            ValidateFolder(newFilesfolderPath);

            XLWorkbook CurrentWorkBook = new XLWorkbook(path, XLEventTracking.Disabled);
            string filename = string.IsNullOrEmpty(template_name) ? Path.GetFileNameWithoutExtension(DefaultFileName) : Path.GetFileNameWithoutExtension(template_name);
            string extension = string.IsNullOrEmpty(template_name) ? Path.GetExtension(DefaultFileName) : Path.GetExtension(template_name);

            if (!IsValidExtension(extension))
                throw new Exception(string.Format("The extension {0} is not valid, valid extensions:{1}", extension, string.Join(",", SupportedExtensions)));

            IXLWorksheet workSheet = sheetName == null ? CurrentWorkBook.Worksheet(1) : CurrentWorkBook.Worksheet(sheetName);
            IEnumerable<IXLRow> rows = workSheet.RowsUsed(XLCellsUsedOptions.Contents, r => r.RowNumber() > 1);
            IList<string> headers = (customHeaders == null ? workSheet.FirstRowUsed(XLCellsUsedOptions.Contents).CellsUsed(XLCellsUsedOptions.Contents).Select(c => c.GetString()) : customHeaders).ToList();

            int fileRowsNumber = workSheet.LastRowUsed(XLCellsUsedOptions.Contents).RowNumber() - 1;
            int remainderRow = fileRowsNumber % rowsQuantity;
            int adittionalRowsPerSheet = remainderRow > 0 ? Convert.ToInt32(Math.Ceiling((double)remainderRow / fileRowsNumber)) : 0;
            int filesQuantity = (fileRowsNumber / rowsQuantity) + adittionalRowsPerSheet;
            int iterations = 0;
            int filesIndexName = 1;

            XLWorkbook newWorbook = new XLWorkbook(XLEventTracking.Disabled);
            IXLWorksheet newWorksheet = newWorbook.Worksheets.Add(DefaultSheetName);
            string fullFileName = Path.Combine(newFilesfolderPath, filename + filesIndexName.ToString() + extension);

            SetHeaders(newWorbook, headers, newWorksheet.Name);

            while (iterations < filesQuantity)
            {
                IEnumerable<IXLRow> currentRows = rows.Skip((rowsQuantity * iterations)).Take(rowsQuantity);
                ICollection<object[]> currentCells = new List<object[]>();

                foreach (IXLRow item in currentRows)
                {
                    currentCells.Add(item.CellsUsed(XLCellsUsedOptions.Contents).Select(c => c.Value).ToArray());
                }

                newWorksheet.Cell(2, 1).InsertData(currentCells);

                Save(newWorbook, fullFileName);
                newWorbook = new XLWorkbook(XLEventTracking.Disabled);
                newWorksheet = newWorbook.Worksheets.Add(DefaultSheetName);
                SetHeaders(newWorbook, headers, newWorksheet.Name);

                filesIndexName++;
                iterations++;

                fullFileName = Path.Combine(newFilesfolderPath, filename + filesIndexName.ToString() + extension);
            }
        }

        /// <summary>Join an amount of excel files into one, if excel sheet is full then create a new one an continue joining</summary>
        /// <param name="filesFolderPath">Folder to use to get the files to join</param>
        /// <param name="newFilePath">Path of the new Excel file incliding the name</param>
        /// <param name="sheetName">Name for the new file sheet</param>
        /// <param name="filesSheetNameToUse">If files have the same sheetName then it will look for it, if not let this param null to take the first one of each</param>
        /// <param name="customHeaders">Titles to use in the new Excel file. If is null it will use the first found file headers</param>
        public static void Join(string filesFolderPath, string newFilePath, string sheetName = null, string filesSheetNameToUse = null, IEnumerable<string> customHeaders = null)
        {
            ValidateFolder(filesFolderPath);

            XLWorkbook newWorkBook = new XLWorkbook(XLEventTracking.Disabled);
            IXLWorksheet workSheet = sheetName == null ? newWorkBook.Worksheets.Add(DefaultSheetName) : newWorkBook.Worksheets.Add(sheetName);
            string currentSheetName = workSheet.Name;
            bool hasHeader = false;

            foreach (string filePath in Directory.GetFiles(filesFolderPath))
            {
                if (!IsValidExtension(Path.GetExtension(filePath)))
                    continue;

                XLWorkbook fileWorkbook = new XLWorkbook(filePath, XLEventTracking.Disabled);
                IXLWorksheet fileWorkSheet = filesSheetNameToUse == null ? fileWorkbook.Worksheet(1) : fileWorkbook.Worksheet(filesSheetNameToUse);
                IXLRows rows = fileWorkSheet.RowsUsed(XLCellsUsedOptions.Contents, r => r.RowNumber() > 1);

                if (!hasHeader)
                {
                    SetHeaders(newWorkBook, customHeaders ?? fileWorkSheet.FirstRowUsed(XLCellsUsedOptions.Contents).CellsUsed(XLCellsUsedOptions.Contents).Select(c => c.GetString()));
                    hasHeader = true;
                }

                ICollection<object[]> currentCells = new List<object[]>();
                foreach (IXLRow item in rows)
                {
                    currentCells.Add(item.CellsUsed(XLCellsUsedOptions.Contents).Select(c => c.Value).ToArray());
                }

                workSheet.Cell(2, 1).InsertData(currentCells);
                workSheet = PrepareNewSheet(newWorkBook, workSheet);
            }

            newWorkBook.Worksheets.Delete(workSheet.Name);
            Save(newWorkBook, newFilePath);
        }

        /// <summary>Save the current excel file changes</summary>
        public void Save()
        {
            SaveOptions so = new SaveOptions()
            {
                GenerateCalculationChain = false,
                EvaluateFormulasBeforeSaving = false,
                ValidatePackage = false,
            };
            Workbook.Save(so);
        }

        /// <summary>Save the new excel file</summary>
        /// <param name="path">Path to use to save the new excel file</param>
        public void Save(string path)
        {
            SaveOptions so = new SaveOptions()
            {
                GenerateCalculationChain = false,
                EvaluateFormulasBeforeSaving = false,
                ValidatePackage = false,
            };
            Workbook.SaveAs(path, so);
        }

        private static void Save(XLWorkbook workbook, string path)
        {
            SaveOptions so = new SaveOptions()
            {
                GenerateCalculationChain = false,
                EvaluateFormulasBeforeSaving = false,
                ValidatePackage = false,
            };
            workbook.SaveAs(path, so);
        }

        private static void RemoveTableStyles(IXLTable table)
        {
            table.ShowAutoFilter = false;
            table.Theme = XLTableTheme.None;
        }

        private static bool IsValidExtension(string extension)
        {
            return SupportedExtensions.Contains(Path.GetExtension(extension));
        }

        private static T GetObjectFromRow<T>(IList<string> headers, IXLRow row) where T : new()
        {
            Type type = typeof(T);
            T newObj = new T();

            for (int i = 0; i < headers.Count; i++)
            {
                if (HasProperty(newObj, headers[i]))
                {
                    PropertyInfo property = type.GetProperty(headers[i]);
                    if (IsValidPropertyType(property.PropertyType))
                    {
                        property.SetValue(newObj, Convert.ChangeType(row.Cell(i + 1).Value, property.PropertyType), null);
                    }
                }
            }
            return newObj;
        }

        private static bool HasProperty(object obj, string name)
        {
            try
            {
                return obj.GetType().GetProperty(name) != null;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsValidPropertyType(Type propType)
        {
            if (propType.IsValueType)
                return true;

            return ValidConvertedValueTypes.Contains(propType);
        }

        private static IXLWorksheet AddNewSheet(XLWorkbook workbook)
        {
            IXLWorksheet newWorksheet = null;
            List<string> sheets = workbook.Worksheets.Select(w => w.Name).ToList();
            string newSheetName = DefaultSheetName.Substring(0, DefaultSheetName.Length - 1) + (sheets.Count + 1);

            if (newSheetName.Length > 31)
                newSheetName = newSheetName.Substring(0, 31);

            if (sheets.Contains(newSheetName))
                newSheetName = DateTime.Now.ToString("MM-dd-yyyy h:mm tt");

            newWorksheet = workbook.Worksheets.Add(newSheetName);
            return newWorksheet;
        }

        private static int AddHeaders(IXLWorksheet workSheet, IEnumerable<string> headers)
        {
            int rowIndex = 1;
            if (headers != null)
            {
                int cellIndex = 1;
                IXLRow row = workSheet.FirstRow();
                foreach (string value in headers)
                {
                    row.Cell(cellIndex).Value = value;
                    cellIndex++;
                }
                rowIndex++;
            }
            return rowIndex;
        }

        private static int AddHeaders(IXLWorksheet workSheet, System.Data.DataColumnCollection headers)
        {
            int rowIndex = 1;
            if (headers != null)
            {
                int cellIndex = 1;
                IXLRow row = workSheet.FirstRow();
                foreach (System.Data.DataColumn value in headers)
                {
                    row.Cell(cellIndex).Value = value.ColumnName;
                    cellIndex++;
                }
                rowIndex++;
            }
            return rowIndex;
        }

        private static void ValidateFolder(string folder)
        {
            if (!Directory.Exists(folder))
                throw new Exception(string.Format("The directory {0} is not valid", folder));
        }

        private static void ValidateFile(string path)
        {
            if (!File.Exists(path))
                throw new Exception(string.Format("The Filepath {0} is not valid", path));
        }

        private static IXLWorksheet PrepareNewSheet(XLWorkbook Workbook, IXLWorksheet workSheet)
        {
            IEnumerable<string> headers = workSheet.FirstRow().CellsUsed(XLCellsUsedOptions.Contents).Select(c => c.GetString());
            workSheet = AddNewSheet(Workbook);
            AddHeaders(workSheet, headers);
            return workSheet;
        }

        private void IsValidLimitPerSheet()
        {
            if (LimitPerSheet > MaxRowLimit)
                throw new Exception(string.Format("The row limit cannot be greater than {0}", MaxRowLimit));
        }
    }

    ///<summary>And enum direction to set movemment</summary>
    public enum Direction
    {
        BOTTOM = 0,
        RIGHT = 1,
        TOP = 2,
        LEFT = 3,
        RIGHTUP = 4,
        RIGHTDOWN = 5,
        LEFTUP = 6,
        LEFTDOWN = 7
    }

    ///<summary>An enum of keys to use</summary>
    public enum KEYCODE
    {

        // Summary:
        //     BACKSPACE key
        BACKSPACE = 8,
        //
        // Summary:
        //     TAB key
        TAB = 9,
        //
        // Summary:
        //     CLEAR key
        CLEAR = 12,
        //
        // Summary:
        //     ENTER key
        ENTER = 13,
        //
        // Summary:
        //     SHIFT key
        SHIFT = 16,
        //
        // Summary:
        //     CTRL key
        CONTROL = 17,
        //
        // Summary:
        //     ALT key
        ALT = 18,
        //
        // Summary:
        //     PAUSE key
        PAUSE = 19,
        //
        // Summary:
        //     CAPS LOCK key
        CAPITAL = 20,
        //
        // Summary:
        //     Input Method Editor (IME) Kana mode
        KANA = 21,
        //
        // Summary:
        //     IME Hanguel mode (maintained for compatibility; use HANGUL)
        HANGEUL = 21,
        //
        // Summary:
        //     IME Hangul mode
        HANGUL = 21,
        //
        // Summary:
        //     IME Junja mode
        JUNJA = 23,
        //
        // Summary:
        //     IME final mode
        FINAL = 24,
        //
        // Summary:
        //     IME Kanji mode
        KANJI = 25,
        //
        // Summary:
        //     IME Hanja mode
        HANJA = 25,
        //
        // Summary:
        //     ESC key
        ESCAPE = 27,
        //
        // Summary:
        //     IME convert
        CONVERT = 28,
        //
        // Summary:
        //     IME nonconvert
        NONCONVERT = 29,
        //
        // Summary:
        //     IME accept
        ACCEPT = 30,
        //
        // Summary:
        //     IME mode change request
        MODECHANGE = 31,
        //
        // Summary:
        //     SPACEBAR
        SPACE = 32,
        //
        // Summary:
        //     PAGE UP key
        PAGE_UP = 33,
        //
        // Summary:
        //     PAGE DOWN key
        PAGE_DOWN = 34,
        //
        // Summary:
        //     END key
        END = 35,
        //
        // Summary:
        //     HOME key
        HOME = 36,
        //
        // Summary:
        //     LEFT ARROW key
        LEFT = 37,
        //
        // Summary:
        //     UP ARROW key
        UP = 38,
        //
        // Summary:
        //     RIGHT ARROW key
        RIGHT = 39,
        //
        // Summary:
        //     DOWN ARROW key
        DOWN = 40,
        //
        // Summary:
        //     SELECT key
        SELECT = 41,
        //
        // Summary:
        //     PRINT key
        PRINT = 42,
        //
        // Summary:
        //     EXECUTE key
        EXECUTE = 43,
        //
        // Summary:
        //     PRINT SCREEN key
        SNAPSHOT = 44,
        //
        // Summary:
        //     INS key
        INSERT = 45,
        //
        // Summary:
        //     DEL key
        DELETE = 46,
        //
        // Summary:
        //     HELP key
        HELP = 47,
        //
        // Summary:
        //     0 key
        D0 = 48,
        //
        // Summary:
        //     1 key
        D1 = 49,
        //
        // Summary:
        //     2 key
        D2 = 50,
        //
        // Summary:
        //     3 key
        D3 = 51,
        //
        // Summary:
        //     4 key
        D4 = 52,
        //
        // Summary:
        //     5 key
        D5 = 53,
        //
        // Summary:
        //     6 key
        D6 = 54,
        //
        // Summary:
        //     7 key
        D7 = 55,
        //
        // Summary:
        //     8 key
        D8 = 56,
        //
        // Summary:
        //     9 key
        D9 = 57,
        //
        // Summary:
        //     A key
        A = 65,
        //
        // Summary:
        //     B key
        B = 66,
        //
        // Summary:
        //     C key
        C = 67,
        //
        // Summary:
        //     D key
        D = 68,
        //
        // Summary:
        //     E key
        E = 69,
        //
        // Summary:
        //     F key
        F = 70,
        //
        // Summary:
        //     G key
        G = 71,
        //
        // Summary:
        //     H key
        H = 72,
        //
        // Summary:
        //     I key
        I = 73,
        //
        // Summary:
        //     J key
        J = 74,
        //
        // Summary:
        //     K key
        K = 75,
        //
        // Summary:
        //     L key
        L = 76,
        //
        // Summary:
        //     M key
        M = 77,
        //
        // Summary:
        //     N key
        N = 78,
        //
        // Summary:
        //     O key
        O = 79,
        //
        // Summary:
        //     P key
        P = 80,
        //
        // Summary:
        //     Q key
        Q = 81,
        //
        // Summary:
        //     R key
        R = 82,
        //
        // Summary:
        //     S key
        S = 83,
        //
        // Summary:
        //     T key
        T = 84,
        //
        // Summary:
        //     U key
        U = 85,
        //
        // Summary:
        //     V key
        V = 86,
        //
        // Summary:
        //     W key
        W = 87,
        //
        // Summary:
        //     X key
        X = 88,
        //
        // Summary:
        //     Y key
        Y = 89,
        //
        // Summary:
        //     Z key
        Z = 90,
        //
        // Summary:
        //     Left Windows key (Microsoft Natural keyboard)
        L_WIN = 91,
        //
        // Summary:
        //     Right Windows key (Natural keyboard)
        R_WIN = 92,
        //
        // Summary:
        //     Applications key (Natural keyboard)
        APPS = 93,
        //
        // Summary:
        //     Computer Sleep key
        SLEEP = 95,
        //
        // Summary:
        //     Numeric keypad 0 key
        NUMPAD_0 = 96,
        //
        // Summary:
        //     Numeric keypad 1 key
        NUMPAD_1 = 97,
        //
        // Summary:
        //     Numeric keypad 2 key
        NUMPAD_2 = 98,
        //
        // Summary:
        //     Numeric keypad 3 key
        NUMPAD_3 = 99,
        //
        // Summary:
        //     Numeric keypad 4 key
        NUMPAD_4 = 100,
        //
        // Summary:
        //     Numeric keypad 5 key
        NUMPAD_5 = 101,
        //
        // Summary:
        //     Numeric keypad 6 key
        NUMPAD_6 = 102,
        //
        // Summary:
        //     Numeric keypad 7 key
        NUMPAD_7 = 103,
        //
        // Summary:
        //     Numeric keypad 8 key
        NUMPAD_8 = 104,
        //
        // Summary:
        //     Numeric keypad 9 key
        NUMPAD_9 = 105,
        //
        // Summary:
        //     Multiply key
        MULTIPLY = 106,
        //
        // Summary:
        //     Add key
        ADD = 107,
        //
        // Summary:
        //     Separator key
        SEPARATOR = 108,
        //
        // Summary:
        //     Subtract key
        SUBTRACT = 109,
        //
        // Summary:
        //     Decimal key
        DECIMAL = 110,
        //
        // Summary:
        //     Divide key
        DIVIDE = 111,
        //
        // Summary:
        //     F1 key
        F1 = 112,
        //
        // Summary:
        //     F2 key
        F2 = 113,
        //
        // Summary:
        //     F3 key
        F3 = 114,
        //
        // Summary:
        //     F4 key
        F4 = 115,
        //
        // Summary:
        //     F5 key
        F5 = 116,
        //
        // Summary:
        //     F6 key
        F6 = 117,
        //
        // Summary:
        //     F7 key
        F7 = 118,
        //
        // Summary:
        //     F8 key
        F8 = 119,
        //
        // Summary:
        //     F9 key
        F9 = 120,
        //
        // Summary:
        //     F10 key
        F10 = 121,
        //
        // Summary:
        //     F11 key
        F11 = 122,
        //
        // Summary:
        //     F12 key
        F12 = 123,
        //
        // Summary:
        //     F13 key
        F13 = 124,
        //
        // Summary:
        //     F14 key
        F14 = 125,
        //
        // Summary:
        //     F15 key
        F15 = 126,
        //
        // Summary:
        //     F16 key
        F16 = 127,
        //
        // Summary:
        //     F17 key
        F17 = 128,
        //
        // Summary:
        //     F18 key
        F18 = 129,
        //
        // Summary:
        //     F19 key
        F19 = 130,
        //
        // Summary:
        //     F20 key
        F20 = 131,
        //
        // Summary:
        //     F21 key
        F21 = 132,
        //
        // Summary:
        //     F22 key
        F22 = 133,
        //
        // Summary:
        //     F23 key
        F23 = 134,
        //
        // Summary:
        //     F24 key
        F24 = 135,
        //
        // Summary:
        //     NUM LOCK key
        NUMLOCK = 144,
        //
        // Summary:
        //     SCROLL LOCK key
        SCROLL = 145,
        //
        // Summary:
        //     Left SHIFT key - Used only as parameters to GetAsyncKEYState() and GetKeyState()
        L_SHIFT = 160,
        //
        // Summary:
        //     Right SHIFT key - Used only as parameters to GetAsyncKeyState() and GetKeyState()
        R_SHIFT = 161,
        //
        // Summary:
        //     Left CONTROL key - Used only as parameters to GetAsyncKeyState() and GetKeyState()
        L_CTRL = 162,
        //
        // Summary:
        //     Right CONTROL key - Used only as parameters to GetAsyncKeyState() and GetKeyState()
        R_CTRL = 163,
        //
        // Summary:
        //     Left MENU key - Used only as parameters to GetAsyncKeyState() and GetKeyState()
        L_ALT = 164,
        //
        // Summary:
        //     Right MENU key - Used only as parameters to GetAsyncKeyState() and GetKeyState()
        R_ALT = 165,
        //
        // Summary:
        //     Windows 2000/XP: Browser Back key
        BROWSER_BACK = 166,
        //
        // Summary:
        //     Windows 2000/XP: Browser Forward key
        BROWSER_FORWARD = 167,
        //
        // Summary:
        //     Windows 2000/XP: Browser Refresh key
        BROWSER_REFRESH = 168,
        //
        // Summary:
        //     Windows 2000/XP: Browser Stop key
        BROWSER_STOP = 169,
        //
        // Summary:
        //     Windows 2000/XP: Browser Search key
        BROWSER_SEARCH = 170,
        //
        // Summary:
        //     Windows 2000/XP: Browser Favorites key
        BROWSER_FAVORITES = 171,
        //
        // Summary:
        //     Windows 2000/XP: Browser Start and Home key
        BROWSER_HOME = 172,
        //
        // Summary:
        //     Windows 2000/XP: Volume Mute key
        VOLUME_MUTE = 173,
        //
        // Summary:
        //     Windows 2000/XP: Volume Down key
        VOLUME_DOWN = 174,
        //
        // Summary:
        //     Windows 2000/XP: Volume Up key
        VOLUME_UP = 175,
        //
        // Summary:
        //     Windows 2000/XP: Next Track key
        MEDIA_NEXT_TRACK = 176,
        //
        // Summary:
        //     Windows 2000/XP: Previous Track key
        MEDIA_PREV_TRACK = 177,
        //
        // Summary:
        //     Windows 2000/XP: Stop Media key
        MEDIA_STOP = 178,
        //
        // Summary:
        //     Windows 2000/XP: Play/Pause Media key
        MEDIA_PLAY_PAUSE = 179,
        //
        // Summary:
        //     Windows 2000/XP: Start Mail key
        LAUNCH_MAIL = 180,
        //
        // Summary:
        //     Windows 2000/XP: Select Media key
        LAUNCH_MEDIA_SELECT = 181,
        //
        // Summary:
        //     Windows 2000/XP: Start Application 1 key
        LAUNCH_APP1 = 182,
        //
        // Summary:
        //     Windows 2000/XP: Start Application 2 key
        LAUNCH_APP2 = 183,
        //
        // Summary:
        //     Used for miscellaneous characters; it can vary by keyboard. Windows 2000/XP:
        //     For the US standard keyboard, the ';:' key
        OEM_1 = 186,
        //
        // Summary:
        //     Windows 2000/XP: For any country/region, the '+' key
        OEM_PLUS = 187,
        //
        // Summary:
        //     Windows 2000/XP: For any country/region, the ',' key
        OEM_COMMA = 188,
        //
        // Summary:
        //     Windows 2000/XP: For any country/region, the '-' key
        OEM_MINUS = 189,
        //
        // Summary:
        //     Windows 2000/XP: For any country/region, the '.' key
        OEM_PERIOD = 190,
        //
        // Summary:
        //     Used for miscellaneous characters; it can vary by keyboard. Windows 2000/XP:
        //     For the US standard keyboard, the '/?' key
        OEM_2 = 191,
        //
        // Summary:
        //     Used for miscellaneous characters; it can vary by keyboard. Windows 2000/XP:
        //     For the US standard keyboard, the '`~' key
        OEM_3 = 192,
        //
        // Summary:
        //     Used for miscellaneous characters; it can vary by keyboard. Windows 2000/XP:
        //     For the US standard keyboard, the '[{' key
        OEM_4 = 219,
        //
        // Summary:
        //     Used for miscellaneous characters; it can vary by keyboard. Windows 2000/XP:
        //     For the US standard keyboard, the '\|' key
        OEM_5 = 220,
        //
        // Summary:
        //     Used for miscellaneous characters; it can vary by keyboard. Windows 2000/XP:
        //     For the US standard keyboard, the ']}' key
        OEM_6 = 221,
        //
        // Summary:
        //     Used for miscellaneous characters; it can vary by keyboard. Windows 2000/XP:
        //     For the US standard keyboard, the 'single-quote/double-quote' key
        OEM_7 = 222,
        //
        // Summary:
        //     Used for miscellaneous characters; it can vary by keyboard.
        OEM_8 = 223,
        //
        // Summary:
        //     Windows 2000/XP: Either the angle bracket key or the backslash key on the
        //     RT 102-key keyboard
        OEM_102 = 226,
        //
        // Summary:
        //     Windows 95/98/Me, Windows NT 4.0, Windows 2000/XP: IME PROCESS key
        PROCESSKEY = 229,
        //
        // Summary:
        //     Windows 2000/XP: Used to pass Unicode characters as if they were keystrokes.
        //     The PACKET key is the low word of a 32-bit Virtual Key value used for non-keyboard
        //     input methods. For more information, see Remark in KEYBDINPUT, SendInput,
        //     WM_KEYDOWN, and WM_KEYUP
        PACKET = 231,
        //
        // Summary:
        //     Attn key
        ATTN = 246,
        //
        // Summary:
        //     CrSel key
        CRSEL = 247,
        //
        // Summary:
        //     ExSel key
        EXSEL = 248,
        //
        // Summary:
        //     Erase EOF key
        EREOF = 249,
        //
        // Summary:
        //     Play key
        PLAY = 250,
        //
        // Summary:
        //     Zoom key
        ZOOM = 251,
        //
        // Summary:
        //     Reserved
        NONAME = 252,
        //
        // Summary:
        //     PA1 key
        PA1 = 253,
        //
        // Summary:
        //     Clear key
        OEM_CLEAR = 254,
    }

    public class Helpers
    {
        public static void Wait(int milliseconds)
        {
            System.Threading.Thread.Sleep(milliseconds);
        }

        public static void WaitSecond()
        {
            System.Threading.Thread.Sleep(1000);
        }
    }
}
