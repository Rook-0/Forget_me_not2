namespace AbandonedCalls2
{
    public class Calls
    {
        public string WhenCalled { get; set; }
        public string Importance { get; set; }
        public string Tta { get; set; }
        public string ContactNum { get; set; }
        public string Status { get; set; }
        //public bool Active { get; set; }

        public Calls(string wc, string Imp, string TimTA, string cn, string st)
        {
            WhenCalled = wc;
            Importance = Imp;
            Tta = TimTA;
            ContactNum = cn;
            Status = st;
        }
    }
}
