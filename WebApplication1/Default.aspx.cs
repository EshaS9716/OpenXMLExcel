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
        }

        protected void LinkButton1_Click(object sender, EventArgs e)
        {
            GetDataFromAPI.LoadAPI();
            Export_Data.WriteToExcel();

            Response.ContentType = "Application/x-msexcel";
            Response.AppendHeader("Content-Disposition", "attachment; filename=Rocket Product Hierarchy.xlsx");
            Response.TransmitFile(DesPath);
            Response.End();
        }
    }
}