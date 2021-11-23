using System;
using System.Web.UI;
using Ionic.Zip;

namespace ExcelAppOpenXML
{
    public partial class _Default : Page
    {
        public static string SourcePath { get; set; }
        public static string TierSourcePath { get; set; }
        public static string DesPath { get; set; }
        public static string TierDesPath { get; set; }
        public static bool WasImportSuccessful { get; set; }

        protected void Page_Load(object sender, EventArgs e)
        {
            SourcePath = Server.MapPath("~/Template/MyDataTemplate.xlsx");
            DesPath = Server.MapPath("~/Template/Rocket Product Hierarchy.xlsx");

            TierSourcePath = Server.MapPath("~/Template/TierDataTemplate.xlsx");
            TierDesPath = Server.MapPath("~/Template/Tier Hierarchy.xlsx");

            this.Title = "Downloading Excel...";
            DownLoadExcel();
        }

        protected void DownLoadExcel()
        {
            WasImportSuccessful = false;
            if (!GetDataFromAPI.LoadAPI())
            {
                Export_Data.WriteToExcel();
            }
            if (WasImportSuccessful)
            {
                GetDataFromAPI.CopyToMaster();
            }
            GetDataFromAPI.DisposeUsedResources();

            using (ZipFile zip = new ZipFile())
            {
                zip.AlternateEncodingUsage = ZipOption.AsNecessary;
                zip.AddFile(DesPath, ""); 
                zip.AddFile(TierDesPath, "");

                Response.Clear();
                Response.BufferOutput = false;
                string zipName = String.Format("Hierarchy Data ({0}).zip", Export_Data.date);
                Response.ContentType = "application/zip";
                Response.AddHeader("content-disposition", "attachment; filename=" + zipName);
                zip.Save(Response.OutputStream);
                Response.End();
            }
        }
    }
}