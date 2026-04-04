namespace PainTrax.Web.Helper
{
    public class AppConstant
    {
        public static IWebHostEnvironment environment { get; set; }

        

        public static void WriteLogFile(string str)
        {

            try
            {

                if (!String.IsNullOrEmpty(str))
                {
                    try
                    {
                        string filename = DateTime.Now.ToString("MMMddyyyy");

                        string todaysfilepath = environment.WebRootPath + @"/Logfiles/Log_" + filename + ".txt";//System.Web.HttpContext.Current.Server.MapPath(@"~/Logfiles/Log_" + filename + ".txt");

                        if (!Directory.Exists(environment.WebRootPath + @"/Logfiles/"))
                        {
                            Directory.CreateDirectory(environment.WebRootPath + @"/Logfiles/");
                        }

                        if (!File.Exists(todaysfilepath))
                        {
                            File.Create(todaysfilepath).Close();
                        }
                        using (StreamWriter sw = File.AppendText(todaysfilepath))
                        {
                            sw.WriteLine("#");
                            sw.WriteLine(DateTime.Now.ToString());
                            sw.WriteLine(str);
                        }
                    }
                    catch (Exception ex)
                    {
                        //    AppConstant.WriteLogFile(Convert.ToString(ex));
                    }
                }
            }
            catch
            {

                throw;
            }

        }

    }
}
