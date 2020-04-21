using Quellatalo.Nin.TheHands;
using System.Collections.Generic;
using System.Drawing;

namespace CShauto
{
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

}
