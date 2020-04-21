namespace CShauto
{
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
