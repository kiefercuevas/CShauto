namespace CShauto
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
}
