using System;
using System.Net.Mail;
using System.Net.Mime;
using System.Web;
using System.Web.UI;
using Ionic.Zip;

namespace ExcelAppOpenXML
{
    public partial class _Default : Page
    {
        public static string SourcePath { get; set; }
        public static string TierSourcePath { get; set; }
        public static string SKUSourcePath { get; set; }
        public static string DesPath { get; set; }
        public static string TierDesPath { get; set; }
        public static string SKUDesPath { get; set; }
        public static bool WasImportSuccessful { get; set; }
        public static string Email { get; set; }

        protected void Page_Load(object sender, EventArgs e)
        {
            var queryString = HttpContext.Current.Request.QueryString.GetValues(0);
            Email = queryString[0];

            SourcePath = Server.MapPath("~/Template/MyDataTemplate.xlsx");
            DesPath = Server.MapPath("~/Template/Rocket Product Hierarchy_UAT.xlsx");

            TierSourcePath = Server.MapPath("~/Template/TierDataTemplate.xlsx");
            TierDesPath = Server.MapPath("~/Template/Tier Hierarchy_UAT.xlsx");

            SKUSourcePath = Server.MapPath("~/Template/SKUDataTemplate.xlsx");
            SKUDesPath = Server.MapPath("~/Template/SKU Hierarchy_UAT.xlsx");

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
                ContentType ct1 = new ContentType("application/vnd.ms-excel");
                ContentType ct2 = new ContentType("application/vnd.ms-excel");
                ContentType ct3 = new ContentType("application/vnd.ms-excel");

                Attachment attachment1 = new Attachment(DesPath, ct1);
                Attachment attachment2 = new Attachment(TierDesPath, ct2);
                Attachment attachment3 = new Attachment(SKUDesPath, ct3);
                SendAsEmail.SendEmail(Email, attachment1, attachment2, attachment3);
            }
        }
    }
}