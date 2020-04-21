using Quellatalo.Nin.TheEyes;
using Quellatalo.Nin.TheEyes.Pattern.CV.Image;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;

namespace CShauto
{
    /// <summary>A class thar provides some methods to locate images on screen</summary>
    public static class Image
    {

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

        private static Rect FixCoordinates(Rectangle firstRect, Rectangle secondRect)
        {
            return new Rect(firstRect.X + secondRect.X, firstRect.Y + secondRect.Y, secondRect.Width, secondRect.Height);
        }
        /// <summary>A helper to validate an image path</summary>
        private static void ValidateImage(string path)
        {
            if (!File.Exists(path))
                throw new Exception(string.Format("The path {0} is not valid", path));
        }

    }

}
