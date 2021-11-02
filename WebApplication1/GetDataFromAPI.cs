using MySqlConnector;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using RestSharp;
using System;
using System.Data;
using System.Linq;

namespace ExcelAppOpenXML
{
    public class GetDataFromAPI
    {
        public static DataTable dt = new DataTable();
        public static DataTable dataTable1 = new DataTable();
        public static DataTable dataTable2 = new DataTable();
        public static DataTable dataTable3 = new DataTable();
        public static DataTable dataTable4 = new DataTable();

        public static MySql.Data.MySqlClient.MySqlConnection dbConn = new MySql.Data.MySqlClient.MySqlConnection("user id=esahu;server=walstgpimcore01;database=rocket_hierarchy;password=Dev*eSha");
        public static string baseUrl = "http://walprdpimcore01.rocketsoftware.com/api/productmasterlisting";
        public static string headerName = "token";
        public static string headerValue = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ1c2VyX2lkIjoiMDEiLCJleHAiOjMzMTY0NTg3NzgyLCJpc3MiOiJ3YWxwcmRwaW1jb3JlMDEucm9ja2V0c29mdHdhcmUuY29tIiwiaWF0IjoxNjI4NTg3NzgyfQ.xjqnMHUcMkuFnmgVY3y6EOsNsVMyELOQVI1XZA2oRdE";

        public static bool LoadAPI()
        {
            dt = ReadFromApi();
            if (!CompareDbs())
            {
                SendToDb();
                PopulateDatatables();
                return false;
            }
            return true;
        }

        private static bool CompareDbs()
        {
            try
            {
                DataTable tbl = new DataTable();
                dbConn.Open();
                using (MySql.Data.MySqlClient.MySqlCommand cmd = new MySql.Data.MySqlClient.MySqlCommand("APIData", dbConn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.ExecuteNonQuery();
                }

                MySqlConnection connection = new MySqlConnection("user id=esahu;server=walstgpimcore01;database=rocket_hierarchy;password=Dev*eSha;AllowLoadLocalInfile=true");
                connection.Open();
                var bulkCopy = new MySqlBulkCopy(connection);
                bulkCopy.DestinationTableName = "rocket_data_api";
                bulkCopy.WriteToServer(dt);
                connection.Close();

                using (MySql.Data.MySqlClient.MySqlCommand cmd = new MySql.Data.MySqlClient.MySqlCommand("checksum table rocket_hierarchy.rocket_data_pimcore_master, rocket_hierarchy.rocket_data_api", dbConn))
                {
                    tbl.Clear();
                    tbl.Columns.Add("Table");
                    tbl.Columns.Add("Checksum");
                    var mdr = cmd.ExecuteReader();
                    tbl.Load(mdr);
                }

                dbConn.Close();
                return tbl.Rows[0][1].ToString() == tbl.Rows[1][1].ToString();
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

                MySqlConnector.MySqlConnection connection = new MySqlConnector.MySqlConnection("user id=esahu;server=walstgpimcore01;database=rocket_hierarchy;password=Dev*eSha;AllowLoadLocalInfile=true");
                connection.Open();
                var bulkCopy = new MySqlBulkCopy(connection);
                bulkCopy.DestinationTableName = "rocket_data_pimcore";
                bulkCopy.WriteToServer(dt);
                connection.Close();
            }
            catch (Exception ex)
            {
                ErrorLogging.SendErrorToText(ex);
                throw ex;
            }
        }

        public static void PopulateDatatables()
        {
            var pCode = "select * from rocket_pcode_excel";
            var hierarchy = "select * from rocket_hierarchy_excel";
            var pSummary = "select * from product_hierarchy";
            var tSummary = "select * from tier_hierarchy";
            var sp_Page1 = "Page1";
            var sp_Page2 = "Page2";
            var sp_Product = "ProductSummary";
            var sp_Tier = "TierSummary";

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
                using (MySql.Data.MySqlClient.MySqlCommand cmd = new MySql.Data.MySqlClient.MySqlCommand(sp_Product, dbConn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.ExecuteNonQuery();
                }
                using (MySql.Data.MySqlClient.MySqlCommand cmd = new MySql.Data.MySqlClient.MySqlCommand(sp_Tier, dbConn))
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
                using (MySql.Data.MySqlClient.MySqlCommand cmd = new MySql.Data.MySqlClient.MySqlCommand(pSummary, dbConn))
                {
                    var mdr = cmd.ExecuteReader();
                    dataTable3.Load(mdr);
                }
                using (MySql.Data.MySqlClient.MySqlCommand cmd = new MySql.Data.MySqlClient.MySqlCommand(tSummary, dbConn))
                {
                    var mdr = cmd.ExecuteReader();
                    dataTable4.Load(mdr);
                }

                dbConn.Close();
            }
            catch (Exception ex)
            {
                ErrorLogging.SendErrorToText(ex);
                throw ex;
            }
        }

        public static void CopyToMaster()
        {
            var sp_Insert = "CopyToMaster";
            try
            {
                dbConn.Open();
                using (MySql.Data.MySqlClient.MySqlCommand cmd = new MySql.Data.MySqlClient.MySqlCommand(sp_Insert, dbConn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.ExecuteNonQuery();
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