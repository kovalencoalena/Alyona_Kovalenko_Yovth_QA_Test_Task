using System;
using System.Net.Http;
using System.Threading.Tasks;
using System.Linq;
using HtmlAgilityPack;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

class Program
{
    static async Task Main()
    {
        var service = InitGoogleSheets();

        string spreadsheetId = "19Z5Ty2RCKKEq3uDoCehpDRYcJLPeE_70aRPsaXw-n28";
        string range = "Sheet1!A2:C"; // без заголовків

        var request = service.Spreadsheets.Values.Get(spreadsheetId, range);
        var response = await request.ExecuteAsync();

        var rows = response.Values;

        if (rows == null)
        {
            Console.WriteLine("Немає даних");
        }

        var httpClient = new HttpClient();

        for (int i = 0; i < rows.Count; i++)
        {
            var row = rows[i];

            //пропускаемо пусті рядки
            if (row.Count == 0)
            {
                Console.WriteLine($"Row {i+1} empty - skip");
                continue;
            }

            // беру значення з таблиці
            string url = row.Count > 0 ? row[0]?.ToString() ?? "":"";
            string expectedTitle = row.Count > 1 ? row[1]?.ToString() ?? "": "";
            string expectedDesc = row.Count > 2 ? row[2]?.ToString() ?? "" : "";

            //якщо url  пустий  - теж пропускаемо
            if (string.IsNullOrWhiteSpace(url))
            {
                Console.WriteLine($"Row {i+1} no URL - skip");
                continue;
            }

            // отримую реальні дані зі сторінки
            var (actualTitle, actualDesc) = await GetPageData(httpClient, url);

            // порівнюю (ігнорую регістр)
            bool isMatch =
                actualTitle.Trim().ToLower() == expectedTitle.Trim().ToLower() &&
                actualDesc.Trim().ToLower() == expectedDesc.Trim().ToLower();

            Console.WriteLine($"{url} -> {(isMatch ? "OK" : "FAIL")}");

            // фарбую рядок
            await ColorRow(service, spreadsheetId, i + 1, isMatch);
            await Task.Delay(1500);
        }
    }

    // --- підключення до Google ---
    static SheetsService InitGoogleSheets()
    {
        var credential = GoogleCredential
            .FromFile("credentials.json")
            .CreateScoped(SheetsService.Scope.Spreadsheets);
        
         return new SheetsService(new BaseClientService.Initializer()
        {
            HttpClientInitializer = credential,
            ApplicationName = "QA task"
        });
    }

    // --- отримання title + description ---
    static async Task<(string, string)> GetPageData(HttpClient client, string url)
    {
        try
        {
            var html = await client.GetStringAsync(url);

            var doc = new HtmlDocument();
            doc.LoadHtml(html);
            var titleNode = doc.DocumentNode.SelectSingleNode("//title");
            string title = titleNode?.InnerText ?? "";

            var metaNode = doc.DocumentNode
                .SelectNodes("//meta")
                ?.FirstOrDefault(x => x.GetAttributeValue("name", "") == "description");

            string description = metaNode?.GetAttributeValue("content", "") ?? "";

            return (title, description);
        }
        catch
        {
            return ("", "");
        }
    }

    // --- фарбування рядка ---
    static async Task ColorRow(SheetsService service, string spreadsheetId, int rowIndex, bool isMatch)
    {
        var request = new Request
        {
            RepeatCell = new RepeatCellRequest
            {
                Range = new GridRange
                {
                    SheetId = 0,
                    StartRowIndex = rowIndex,
                    EndRowIndex = rowIndex + 1
                },
                Cell = new CellData
                {
                    UserEnteredFormat = new CellFormat
                    {
                      BackgroundColor = new Color
                        {
                            Red = isMatch ? 0 : 1,
                            Green = isMatch ? 1 : 0,
                            Blue = 0
                        }
                    }
                },
                Fields = "userEnteredFormat.backgroundColor"
            }
        };

        var batch = new BatchUpdateSpreadsheetRequest
        {
            Requests = new[] { request }
        };

        await service.Spreadsheets.BatchUpdate(batch, spreadsheetId).ExecuteAsync();
    }
}

        

