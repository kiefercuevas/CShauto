namespace CShauto
{
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
}
