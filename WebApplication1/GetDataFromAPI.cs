using MySqlConnector;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace ExcelAppOpenXML
{
    public class GetDataFromAPI
    {
        public static DataTable DataTable { get; set; }
        public static DataTable dt = new DataTable();
        public static DataTable dataTable1 = new DataTable();
        public static DataTable dataTable2 = new DataTable();
        public static DataTable dataTable3 = new DataTable();

        public static MySql.Data.MySqlClient.MySqlConnection dbConn = new MySql.Data.MySqlClient.MySqlConnection("user id=esahu;server=walstgpimcore01;database=esha_dev;password=Dev*eSha");
        public static string baseUrl = "http://walprdpimcore01.rocketsoftware.com/api/productmasterlisting";
        public static string headerName = "token";
        public static string headerValue = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ1c2VyX2lkIjoiMDEiLCJleHAiOjMzMTY0NTg3NzgyLCJpc3MiOiJ3YWxwcmRwaW1jb3JlMDEucm9ja2V0c29mdHdhcmUuY29tIiwiaWF0IjoxNjI4NTg3NzgyfQ.xjqnMHUcMkuFnmgVY3y6EOsNsVMyELOQVI1XZA2oRdE";

        public static bool LoadAPI()
        {
            ReadFromSP();
            dt = ReadFromApi();
            DataTable = Copy(dt);
            //if (!AreTablesTheSame(DataTable))
            //{
                SendToDb();
                PopulateDatatables();
                return false;
            //}
            //return true;
        }

        static bool AreTablesTheSame(DataTable tbl1)
        {
            try
            {
                DataTable tbl2 = new DataTable();
                using (MySql.Data.MySqlClient.MySqlCommand cmd = new MySql.Data.MySqlClient.MySqlCommand("select * from rocket_data_pimcore", dbConn))
                {
                    dbConn.Open();
                    var mdr = cmd.ExecuteReader();
                    tbl2.Load(mdr);
                    dbConn.Close();
                }
                if (tbl1.Rows.Count != tbl2.Rows.Count || tbl1.Columns.Count != tbl2.Columns.Count)
                    return false;

                for (int i = 0; i < tbl1.Rows.Count; i++)
                {
                    for (int c = 0; c < tbl1.Columns.Count; c++)
                    {
                        if (!Equals(tbl1.Rows[i][c], tbl2.Rows[i][c]))
                            return false;
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                ErrorLogging.SendErrorToText(ex);
                throw ex;
            }
        }

        private static DataTable ReadFromApi()
        {
            var client = new RestClient(baseUrl);
            client.Timeout = -1;
            try
            {
                var request = new RestRequest(Method.GET);
                request.AddHeader(headerName, headerValue);
                IRestResponse response = client.Execute(request);
                return Tabulate(response.Content);
            }
            catch (Exception ex)
            {
                ErrorLogging.SendErrorToText(ex);
                throw ex;
            }
        }

        private static DataTable Tabulate(string json)
        {
            var jsonLinq = JObject.Parse(json);
            // Find the first array using Linq
            var srcArray = jsonLinq.Descendants().Where(d => d is JArray).First();
            var trgArray = new JArray();
            foreach (JObject row in srcArray.Children<JObject>())
            {
                var cleanRow = new JObject();
                foreach (JProperty column in row.Properties())
                {
                    // Only include JValue types
                    if (column.Value is JValue)
                    {
                        cleanRow.Add(column.Name, column.Value);
                    }
                }
                trgArray.Add(cleanRow);
            }
            return JsonConvert.DeserializeObject<DataTable>(trgArray.ToString());

        }

        private static DataTable Copy(DataTable dt)
        {
            var result = from row in dt.AsEnumerable()
                         orderby row.Field<string>("businessUnit"), row.Field<string>("businessUnitGroup"),
                         row.Field<string>("productFamily"), row.Field<string>("productGroup")
                         group row by new
                         {
                             bu = row.Field<string>("businessUnit"),
                             bu_group = row.Field<string>("businessUnitGroup"),
                             bu_product_family = row.Field<string>("productFamily"),
                             bu_product_group = row.Field<string>("productGroup"),
                             product_id = row.Field<string>("productCode"),
                             rocket_product_name = row.Field<string>("productName"),
                             engineer_lead = row.Field<string>("owner"),
                             project_manager = row.Field<string>("productOwnerId"),
                             finance_manager = row.Field<string>("financeOwnerId")
                         } into bu
                         select bu.First();
            DataTable newDataTable = result.CopyToDataTable();
            DataView view = new DataView(newDataTable);
            DataTable selected =
                    view.ToTable("Selected", false, "businessUnit", "businessUnitGroup", "productFamily",
                    "productGroup", "productCode", "productName", "owner", "productOwnerId", "financeOwnerId");
            return selected;
        }

        private static void ReadFromSP()
        {
            try
            {
                using (MySql.Data.MySqlClient.MySqlCommand cmd = new MySql.Data.MySqlClient.MySqlCommand("Page5", dbConn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    dbConn.Open();
                    using (MySql.Data.MySqlClient.MySqlDataAdapter da = new MySql.Data.MySqlClient.MySqlDataAdapter(cmd))
                    {
                        da.Fill(dataTable3);
                    }
                    dbConn.Close();
                }
            }
            catch (Exception ex)
            {
                ErrorLogging.SendErrorToText(ex);
                throw ex;
            }
        }

        private static void SendToDb()
        {
            try
            {
                using (MySql.Data.MySqlClient.MySqlCommand cmd = new MySql.Data.MySqlClient.MySqlCommand("Datasheet", dbConn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    dbConn.Open();
                    cmd.ExecuteNonQuery();
                    dbConn.Close();
                }

                MySqlConnector.MySqlConnection connection = new MySqlConnector.MySqlConnection("user id=esahu;server=walstgpimcore01;database=esha_dev;password=Dev*eSha;AllowLoadLocalInfile=true");
                connection.Open();
                var bulkCopy = new MySqlBulkCopy(connection);
                bulkCopy.DestinationTableName = "rocket_data_pimcore";
                bulkCopy.WriteToServer(DataTable);
                connection.Close();
            }
            catch (Exception ex)
            {
                ErrorLogging.SendErrorToText(ex);
                throw ex;
            }
        }

        private static void PopulateDatatables()
        {
            var pCode = "select * from rocket_pcode_excel";
            var hierarchy = "select * from rocket_hierarchy_excel";
            var sp_Page1 = "Page1";
            var sp_Page2 = "Page2";

            try
            {
                dbConn.Open();

                using (MySql.Data.MySqlClient.MySqlCommand cmd = new MySql.Data.MySqlClient.MySqlCommand(sp_Page1, dbConn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.ExecuteNonQuery();
                }
                using (MySql.Data.MySqlClient.MySqlCommand cmd = new MySql.Data.MySqlClient.MySqlCommand(sp_Page2, dbConn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.ExecuteNonQuery();
                }
                using (MySql.Data.MySqlClient.MySqlCommand cmd = new MySql.Data.MySqlClient.MySqlCommand(hierarchy, dbConn))
                {
                    var mdr = cmd.ExecuteReader();
                    dataTable1.Load(mdr);
                }
                using (MySql.Data.MySqlClient.MySqlCommand cmd = new MySql.Data.MySqlClient.MySqlCommand(pCode, dbConn))
                {
                    var mdr = cmd.ExecuteReader();
                    dataTable2.Load(mdr);
                }

                dbConn.Close();
            }
            catch (Exception ex)
            {
                ErrorLogging.SendErrorToText(ex);
                throw ex;
            }
        }
    }
}