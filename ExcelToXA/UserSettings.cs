namespace ExcelToXA
{
    public class UserSettings
    {
        public string Username { get; set; }
        public string Password { get; set; }
        public Vendor Vendor { get; set; }
        public Environment Environment { get; set; }
        public bool Remember { get; set; }
    }

    public class Vendor
    {
        public string DisplayText { get; set; }
        public int Index { get; set; }
        public int PNColumn {  get; set; }
        public int QTYColumn { get; set; }
        public bool HasHeaders { get; set; }

        public override string ToString()
        {
            return DisplayText;
        }
    }

    public class Environment
    {
        public string DisplayText { get; set; }
        public int Index { get; set; }

        public override string ToString()
        {
            return DisplayText;
        }
    }
}
