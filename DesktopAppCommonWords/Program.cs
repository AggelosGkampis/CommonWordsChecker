using OfficeOpenXml;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;

namespace DesktopAppCommonWords
{
    public static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            ApplicationConfiguration.Initialize();
            Application.Run(new Form1());
        }

        public static DataTable GetCommonWords(ExcelWorkbook workbook)
        {
            var commonWords = new HashSet<string>();
            var worksheet = workbook.Worksheets[0];
            var ignoredWords = new HashSet<string> { "THE", "I", "A", "YOU", "HE", "SHE", "IT", "WE", "THEY", "AN", "AND", "OR",
                                     "BUT", "SO", "IN", "ON", "AT", "IS", "ARE", "WAS", "WERE", "BE", "BEEN", "AM", "DO", "DOES", "DID", "HAVE",
                                     "HAS", "HAD", "OF", "TO", "FOR", "FROM", "BY", "WITH", "THAT", "THIS", "THOSE", "THESE", "IF", "THEN", "ELSE",
                                     "WHEN", "WHERE", "WHILE", "HOW", "WHAT", "WHO", "WHOM", "WHOSE", "NOT", "NO", "YES", "TRUE", "FALSE", "NULL" };

            var column1Words = new Dictionary<string, List<string>>();
            var column2Words = new Dictionary<string, List<string>>();

            // Split words in column 1 and track the row numbers where each word occurs
            for (int i = 1; i <= worksheet.Dimension.Rows; i++)
            {
                var value = worksheet.Cells[i, 1].Value?.ToString()?.Trim()?.ToLower();
                if (string.IsNullOrEmpty(value)) continue;

                var cellAddress = $"A{i}";
                var words = value.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)
                                 .Where(w => !ignoredWords.Contains(w, StringComparer.OrdinalIgnoreCase))
                                 .ToList();

                column1Words[cellAddress] = words;
            }

            // Split words in column 2 and track the row numbers where each word occurs
            for (int i = 1; i <= worksheet.Dimension.Rows; i++)
            {
                var value = worksheet.Cells[i, 2].Value?.ToString()?.Trim()?.ToLower();
                if (string.IsNullOrEmpty(value)) continue;

                var cellAddress = $"B{i}";
                var words = value.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)
                                 .Where(w => !ignoredWords.Contains(w, StringComparer.OrdinalIgnoreCase))
                                 .ToList();

                column2Words[cellAddress] = words;
            }

            DataTable result = new DataTable();
            result.Columns.Add("Common Word");
            result.Columns.Add("Column 1 Matches");
            result.Columns.Add("Column 2 Matches");

            // Find common words between the two columns and track the row numbers where each word occurs
            foreach (var word in column1Words.Values.SelectMany(x => x).Intersect(column2Words.Values.SelectMany(x => x)).Distinct())
            {
                var column1Matches = column1Words.Where(x => x.Value.Contains(word)).Select(x => x.Key);
                var column2Matches = column2Words.Where(x => x.Value.Contains(word)).Select(x => x.Key);

                foreach (var column1Match in column1Matches)
                {
                    foreach (var column2Match in column2Matches)
                    {
                        var row = result.NewRow();
                        row["Common Word"] = word;
                        row["Column 1 Matches"] = worksheet.Cells[column1Match].Value;
                        row["Column 2 Matches"] = worksheet.Cells[column2Match].Value;
                        result.Rows.Add(row);
                    }
                }
            }

            return result;
        }



    }
}