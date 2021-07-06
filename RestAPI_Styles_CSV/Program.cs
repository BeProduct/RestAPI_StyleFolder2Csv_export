using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using CsvHelper;
using OAuth2.Helpers;
using RestSharp;

namespace RestAPI_Styles_CSV
{
    internal class Program
    {
        static readonly string authUrl = "https://id.winks.io/ids";
        static readonly string clientId = "#CLIENT_ID#";
        static readonly string clientSecret = "#SECRET_CODE#";
        static readonly string companyName = "#COMPANY_NAME#";
        static readonly string callbackUrl = "http://localhost:8888/";


        //Add Refresh Token
        private static readonly string refreshToken = "#REFRESH_TOKEN#";

        //Add Style Folder Id
        private static readonly string folderId = "59d35ba9-bc2e-47b6-bd05-9729b14f66a4";

        //Csv File
        private static readonly string csvFilePath = @"d:\temp\csv";
        private static string csvFileName;

        private static object SearchJson
        {
            get
            {
                var searchJson = new
                {
                    filters = new[]
                    {
                        new {field = "status_style", value = (object) "DEVELOPMENT", @operator = "Eq", type = "string"},
                    }
                };
                return searchJson;
            }
        }


        private static void Main(string[] args)
        {

            //Csv File Name
            csvFileName = Path.Combine(csvFilePath, $"{DateTime.Now:yyyy-dd-M--HH-mm-ss}_{folderId}.csv");

            //Get access token
            Console.WriteLine($"Getting access token...");
            var accessToken = Auth.RefreshAccessToken(authUrl, clientId, clientSecret, refreshToken);
            var client = new RestClient("https://developers.beproduct.com/");

            //Pagination
            var pageSize = 200;
            var totalStyleCount = GetTotalStyleCount(folderId, accessToken, SearchJson, client, out var total);
            var totalPage = (total + pageSize - 1) / pageSize;
            Console.WriteLine($"{totalStyleCount} style found!");

            //Add Style Header Fields to Data Dictionary
            var records = GetStyleHeader(totalPage, folderId, pageSize, accessToken, client);

            //Generate Csv File
            Console.WriteLine($"Generating csv file...");
            GenerateCsvFile(records, csvFileName);

            Console.WriteLine($"{csvFileName}");
            Console.WriteLine($"Done!");
            Console.WriteLine("Press any key...");
            Console.ReadKey();
        }

        private static List<Dictionary<string, object>> GetStyleHeader(dynamic totalPage, string folderId, int pageSize, string accessToken,
            RestClient client)
        {
            Console.WriteLine();
            var records = new List<Dictionary<string, object>>();
            for (var pageNumber = 0; pageNumber < totalPage; pageNumber++)
            {
                var request = new RestRequest(
                    $"/api/{companyName}/Style/Headers?folderId={folderId}&pageSize={pageSize}&pageNumber={pageNumber}",
                    Method.POST);
                request.AddHeader("Authorization", "Bearer " + accessToken);
                request.RequestFormat = DataFormat.Json;
                request.AddJsonBody(SearchJson);
                var response = client.Execute<dynamic>(request);

                //
                List<dynamic> fields = response.Data["result"];
                foreach (var field in fields)
                {
                    var rec = new Dictionary<string, object>();

                    //Header Ids
                    rec.Add("folder_id", field["folder"]["id"]);
                    rec.Add("header_id",field["id"]);

                    //Header Custom Fields
                    var headerData = field["headerData"]["fields"];
                    foreach (var data in headerData)
                        if (data["type"] == "PartnerDropDown")
                        {
                            string partnerValue = "";
                            try { partnerValue = Convert.ToString(data["value"]["value"]); }
                            catch { partnerValue = ""; }
                            rec.Add(data["id"], partnerValue);
                        }
                        else
                        {
                            string fieldValue = data["value"] == null ? string.Empty : (string)data["value"];
                            rec.Add(data["id"], fieldValue);
                        }

                    //Colorways
                    var headerColor = field["colorways"];
                    List<dynamic> headerColorCount = headerColor;
                    if (headerColorCount.Count != 0)
                        foreach (KeyValuePair<string, object> color in headerColor[0])
                            if (color.Key == "fields")
                            {
                                //Custom Colorway Fields
                                var customColors = color.Value;
                                foreach (var customColor in (IEnumerable<KeyValuePair<string, object>>) customColors)
                                    rec.Add(customColor.Key, customColor.Value);
                            }
                            else if (color.Key != "id")
                            {
                                //Colorway Fields
                                rec.Add(color.Key, color.Value);
                            }

                    records.Add(rec);
                    Console.Write($"*");
                }
            }
            Console.WriteLine();
            return records;
        }

        private static void GenerateCsvFile(List<Dictionary<string, object>> records, string csvFileName)
        {
            var dt = GetDataTable(records);

            Directory.CreateDirectory(csvFilePath);

            using var writer = new StreamWriter(csvFileName);
            using var csv = new CsvWriter(writer, CultureInfo.InvariantCulture);
            foreach (DataColumn dc in dt.Columns) csv.WriteField(dc.ColumnName);
            csv.NextRecord();
            foreach (DataRow dr in dt.Rows)
            {
                foreach (DataColumn dc in dt.Columns) csv.WriteField(dr[dc]);
                csv.NextRecord();
            }
        }

        private static DataTable GetDataTable(List<Dictionary<string, object>> list)
        {
            var result = new DataTable();
            if (list.Count == 0)
                return result;

            var columnNames = list.SelectMany(dict => dict.Keys).Distinct();
            result.Columns.AddRange(columnNames.Select(c => new DataColumn(c)).ToArray());
            foreach (var item in list)
            {
                var row = result.NewRow();
                foreach (var key in item.Keys) row[key] = item[key];
                result.Rows.Add(row);
            }

            return result;
        }

        private static int GetTotalStyleCount(string folderId, string accessToken, object searchJson,
            RestClient client, out dynamic total)
        {
            var request = new RestRequest(
                $"/api/{companyName}/Style/Headers?folderId={folderId}&pageSize=1&pageNumber=0",
                Method.POST);
            request.AddHeader("Authorization", "Bearer " + accessToken);
            request.RequestFormat = DataFormat.Json;
            request.AddJsonBody(searchJson);
            var response = client.Execute<dynamic>(request);
            total = response.Data["total"];
            return (int) total;
        }
    }
}