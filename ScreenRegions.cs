using System.Drawing;
using System.Windows.Forms;

namespace CShauto
{
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
}
