This project is a C# console application that reads URLs from Google Sheets, extracts page title and meta description, compares them with expected values, and highlights rows based on the result.

Technologies
- C#
- .NET
- HtmlAgilityPack
- Google Sheets API

Setup
1. Install .NET SDK
2. Clone repository
3. Add credentials.json file (not included in repo)
4. Install dependencies:
dotnet restore

Run
dotnet run

Notes
- The script reads data dynamically using header names
- Handles missing meta tags and HTTP errors
- Uses case-insensitive comparison
