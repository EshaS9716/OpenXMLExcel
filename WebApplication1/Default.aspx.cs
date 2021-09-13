using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace ExcelAppOpenXML
{
    public partial class _Default : Page
    {
        public static string SourcePath { get; set; }
        public static string DesPath { get; set; }
        protected void Page_Load(object sender, EventArgs e)
        {
            SourcePath = Server.MapPath("~/Template/MyDataTemplate.xlsx");
            DesPath = Server.MapPath("~/Template/Rocket Product Hierarchy.xlsx");
            this.Title = "Downloading Excel...";
            DownLoadExcel();
        }

        protected void DownLoadExcel()
        {
            if (!GetDataFromAPI.LoadAPI())
            {
                Export_Data.WriteToExcel();
            }

            Response.ContentType = "Application/x-msexcel";
            Response.AddHeader("Content-Disposition", "attachment; filename=" + string.Format("Rocket Product Hierarchy" + " (" + Export_Data.date + ")" +".xlsx"));
            Response.TransmitFile(DesPath);
            Response.End();
        }
    }
}