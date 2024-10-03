namespace ExcelToXA
{
    public class Program
    {
        public static Configuration config;
        public static Main main;
        public static KeyValuePair<string, string> xaLogin = new KeyValuePair<string, string>("", "");
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            ApplicationConfiguration.Initialize();
            Application.Run(new Main());
        }
    }
}