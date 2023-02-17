using OfficeOpenXml;
using System.Collections.Generic;
using System.Data;

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

                var column1 = string.Join(" ", worksheet.Cells["A:A"].Select(c => c.Value?.ToString().Trim().ToLower()));
                var column2 = string.Join(" ", worksheet.Cells["B:B"].Select(c => c.Value?.ToString().Trim().ToLower()));

                var words1 = column1.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                var words2 = column2.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                var ignoredWords = new HashSet<string> { "THE", "I", "A", "YOU", "HE", "SHE", "IT", "WE", "THEY", "AN", "AND", "OR",
                                                         "BUT", "SO", "IN", "ON", "AT", "IS", "ARE", "WAS", "WERE", "BE", "BEEN", "AM", "DO", "DOES", "DID", "HAVE",
                                                         "HAS", "HAD", "OF", "TO", "FOR", "FROM", "BY", "WITH", "THAT", "THIS", "THOSE", "THESE", "IF", "THEN", "ELSE",
                                                         "WHEN", "WHERE", "WHILE", "HOW", "WHAT", "WHO", "WHOM", "WHOSE", "NOT", "NO", "YES", "TRUE", "FALSE", "NULL" };

                foreach (var word1 in words1)
                {
                    if (ignoredWords.Contains(word1, StringComparer.OrdinalIgnoreCase)) continue;
                    if (string.IsNullOrEmpty(word1)) continue;

                    if (words2.Contains(word1))
                    {
                        commonWords.Add(word1);
                    }
                }

                // Create a new DataTable
                var table = new DataTable();

                // Add columns to the DataTable
                table.Columns.Add("Common Word");

                // Loop through the common words and add each row to the DataTable
                foreach (var word in commonWords)
                {
                    table.Rows.Add(word);
                }

                return table;
            }

    }
}
