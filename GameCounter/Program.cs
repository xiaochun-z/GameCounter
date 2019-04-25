using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OfficeOpenXml.Style;

namespace GameCounter
{

    class Program
    {
        static void Main(string[] args)
        {
            var file = new FileInfo("library-04-25-2019.xlsx");

            var xboxoneGames = SpreadsheetReader.ReadGames(file, "Xboxone", 1, 2, 4, 7);
            var xbox360CompatibleGames = SpreadsheetReader.ReadGames(file, "Xbox 360 Compatible", 1, -1, 2, 5);
            var xboxCompatibleGames = SpreadsheetReader.ReadGames(file, "Xbox 1st Gen Compatible", 1, -1, 2, 3);

            var games_Available_On_Xboxone = new List<Game>();
            games_Available_On_Xboxone.AddRange(xboxoneGames);
            games_Available_On_Xboxone.AddRange(xbox360CompatibleGames);
            games_Available_On_Xboxone.AddRange(xboxCompatibleGames);

            var games_Available_On_Ps4 = SpreadsheetReader.ReadGames(file, "PS4", 1, 2, 4, 7);

            var games_on_xboxone_but_unavailable_for_PS4 =
                games_Available_On_Xboxone.Where(a => games_Available_On_Ps4.All(b => b.Name != a.Name)).ToList();

            var games_on_PS4_but_unavailable_for_Xboxone = games_Available_On_Ps4
                .Where(a => games_Available_On_Xboxone.All(b => b.Name != a.Name)).ToList();

            var games_on_both_platforms =
                games_Available_On_Ps4.Where(a => games_Available_On_Xboxone.Any(b => b.Name == a.Name)).ToList();

            Console.WriteLine($"Games available on xboxone: {games_Available_On_Xboxone.Count}");
            Console.WriteLine($"Games available on Ps4: {games_Available_On_Ps4.Count}");
            Console.WriteLine($"Games available on XboxOne but NOT available on PS4: {games_on_xboxone_but_unavailable_for_PS4.Count}");
            Console.WriteLine($"Games available on PS4 but NOT available on XboxOne: {games_on_PS4_but_unavailable_for_Xboxone.Count}");
            Console.WriteLine($"Games available on both platforms: {games_on_both_platforms.Count}");

            // create report.
            Console.WriteLine("creating report...");
            CreateReport(games_on_xboxone_but_unavailable_for_PS4, "microsoft_exclusive.xlsx", "Games on XboxOne or PC");
            CreateReport(games_on_PS4_but_unavailable_for_Xboxone, "PS4_exclusive.xlsx", "Games on PS4 or PC");
            CreateReport(games_on_both_platforms, "PS4_Xboxone.xlsx", "Games on both PS4 and Xboxone");
            Console.WriteLine("done..");
            Console.ReadKey();
        }

        public static void CreateReport(List<Game> games, string file, string sheetName)
        {
            var fileInfo = new FileInfo(file);
            using (var xlPackage = new ExcelPackage())
            {
                var sheet = xlPackage.Workbook.Worksheets.Add(sheetName);
                sheet.Column(1).Width = 58;
                sheet.Column(2).Width = 42;
                sheet.Column(3).Width = 35;
                sheet.Column(4).Width = 11;
                sheet.SetValue(1, 1, "Title");
                sheet.SetValue(1, 2, "Publisher(s)");
                sheet.SetValue(1, 3, "Genre(s)");
                sheet.SetValue(1, 4, "Date");
                sheet.Cells["A1:D1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["A1:D1"].Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
                for (int i = 0; i < games.Count; i++)
                {
                    sheet.SetValue(i + 2, 1, games[i].Name);
                    sheet.SetValue(i + 2, 2, games[i].Publisher);
                    sheet.SetValue(i + 2, 3, games[i].Genre);
                    if (games[i].AddDate.HasValue && games[i].AddDate != DateTime.MinValue)
                        sheet.SetValue(i + 2, 4, games[i].AddDate?.ToString("MM/dd/yyyy"));
                }

                xlPackage.Compression = CompressionLevel.BestSpeed;
                xlPackage.SaveAs(fileInfo);
            }
        }
    }

    public static class SpreadsheetReader
    {
        public static List<Game> ReadGames(FileInfo excel, string sheetName, int nameIndex, int genreIndex, int publisherIndex, int dateIndex)
        {
            var games = new List<Game>();
            using (var xlPackage = new ExcelPackage(excel))
            {
                var ws = xlPackage.Workbook.Worksheets.First(a => a.Name == sheetName);
                for (var i = 2; ; i++)
                {
                    var name = nameIndex > 0 ? ws.GetValue(i, nameIndex) : "N/A";
                    var genre = genreIndex > 0 ? ws.GetValue(i, genreIndex) : "N/A";
                    var publisher = publisherIndex > 0 ? ws.GetValue(i, publisherIndex) : "N/A";
                    var date = dateIndex > 0 ? ws.GetValue(i, dateIndex) : "N/A";

                    if (name == null) break;
                    games.Add(new Game(name?.ToString(), genre?.ToString(), publisher?.ToString(), date?.ToString()));
                }
            }

            return games;
        }
    }

    public class Game
    {
        private readonly DateTime _addDate;

        public Game(string name, string genre, string publisher, string addDate)
        {
            Name = name;
            Genre = genre;
            Publisher = publisher;
            if (!string.IsNullOrWhiteSpace(addDate))
                DateTime.TryParse(addDate, out _addDate);
        }
        public string Name { get; }
        public string Genre { get; }
        public string Publisher { get; }

        public DateTime? AddDate => _addDate;

        public override string ToString()
        {
            return $"{Name}----{Genre}----{Publisher}----{AddDate}";
        }
    }
}