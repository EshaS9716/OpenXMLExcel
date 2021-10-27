using System;
using System.IO;
using context = System.Web.HttpContext;

namespace ExcelAppOpenXML
{
    public static class ErrorLogging
    {
        private static String ErrorlineNo, Errormsg, extype, ErrorLocation;

        public static void SendErrorToText(Exception ex)
        {
            var line = Environment.NewLine;
            int idx = ex.StackTrace.LastIndexOf('\\');
            ErrorlineNo = ex.StackTrace.Substring(idx + 1);
            Errormsg = ex.GetType().Name.ToString();
            extype = ex.GetType().ToString();
            ErrorLocation = ex.Message.ToString();
            _Default.WasImportSuccessful = false;

            try
            {
                string filepath = context.Current.Server.MapPath("~/ExceptionDetailsFile/");  //Text File Path

                if (!Directory.Exists(filepath))
                {
                    Directory.CreateDirectory(filepath);
                }
                filepath = filepath + DateTime.Today.ToString("dd-MM-yy") + ".txt";   //Text File Name
                if (!File.Exists(filepath))
                {
                    File.Create(filepath).Dispose();
                }
                using (StreamWriter sw = File.AppendText(filepath))
                {
                    string error = line 
                        + "Error Line No :" + " " + ErrorlineNo + line
                        + "Error Message:" + " " + Errormsg + line
                        + "Exception Type:" + " " + extype + line
                        + "Error Location :" + " " + ErrorLocation + line;

                    sw.WriteLine("-----------Exception Details on" + " " + DateTime.Now.ToString() + "-----------------");
                    sw.WriteLine("--------------------------------------------------------------------");
                    sw.WriteLine(error);
                    sw.WriteLine("-------------------------------*End*--------------------------------");
                    sw.WriteLine(line);
                    sw.Flush();
                    sw.Close();
                }
            }
            catch (Exception e)
            {
                e.ToString();
            }
        }
    }
}