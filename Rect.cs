namespace CShauto
{
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
}
