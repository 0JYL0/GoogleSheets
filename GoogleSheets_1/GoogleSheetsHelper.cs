using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using System;
using System.Collections.Generic;
using System.IO;

namespace GoogleSheets_1
{
    public class GoogleSheetsHelper
    {
        public SheetsService Service { get; set; }
        const string APPLICATION_NAME = "GroceryStore";
        static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };

        public GoogleSheetsHelper()
        {
            InitializeService();
        }

        private void InitializeService()
        {
            var credential = GetCredentialsFromFile();
            Service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = APPLICATION_NAME
            });
        }

        private GoogleCredential GetCredentialsFromFile()
        {
            GoogleCredential credential;
            using (var stream = new FileStream("secertkey-2a110a4e6a8e.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream).CreateScoped(Scopes);
            }

            return credential;
        }
    }

    //Model for data gathering
    public class Item
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string Category { get; set; }
        public string Price { get; set; }
    }

    //--------
    public static class ItemsMapper
    {
        public static List<Item> MapFromRangeData(IList<IList<object>> values)
        {
            var items = new List<Item>();

            foreach (var value in values)
            {
                Item item = new Item()
                {
                    Id = value[0].ToString(),
                    Name = value[1].ToString(),
                    Category = value[2].ToString(),
                    Price = value[3].ToString()
                };

                items.Add(item);
            }

            return items;
        }

        public static IList<IList<Item>> MapToRangeData(List<Item> item)
        {

            var objectList = new List<Item>();

            foreach(var Itemo in item)
            {
                objectList.Add(Itemo);
            }

            List<IList<Item>>  rangeData = new List<IList<Item>> { objectList };
            Console.WriteLine(rangeData);
          
            return rangeData;
        }
    }
}
