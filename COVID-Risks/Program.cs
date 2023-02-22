namespace COVID_Risks
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            ApplicationConfiguration.Initialize();
            Application.Run(new Form1());


            try
            {
                using (Models.ExcelHelper helper = new Models.ExcelHelper())
                {
                    if (helper.Open(_filePath: Path.Combine(Environment.CurrentDirectory, "testinput.xlsx")))
                    {
                        helper.Set(collumn:"A", row: 1, data: "asdasdasd");

                        helper.Save();
                    }
                }
            }
            catch (Exception e)
            {

                Console.WriteLine(e.Message);
            }
        }
    }
}