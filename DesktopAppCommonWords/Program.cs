using OfficeOpenXml;
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
            var commonWords = new List<string>();
            //var commonWordRowsInColumn1 = new List<string>();
           //var commonWordRowsInColumn2 = new List<string>();
            var worksheet = workbook.Worksheets[0];

            var column1 = string.Join(" ", worksheet.Cells["A:A"].Select(c => c.Value?.ToString().Trim().ToLower()));
            var column2 = string.Join(" ", worksheet.Cells["B:B"].Select(c => c.Value?.ToString().Trim().ToLower()));

            var words1 = column1.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            var words2 = column2.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            var ignoredWords = new List<string> { "THE", "I", "A", "YOU", "HE", "SHE", "IT", "WE", "THEY", "AN", "AND", "OR",
                                                  "BUT", "SO", "IN", "ON", "AT", "IS", "ARE", "WAS", "WERE", "BE", "BEEN", "AM", "DO", "DOES", "DID", "HAVE",
                                                  "HAS", "HAD", "OF", "TO", "FOR", "FROM", "BY", "WITH", "THAT", "THIS", "THOSE", "THESE", "IF", "THEN", "ELSE",
                                                  "WHEN", "WHERE", "WHILE", "HOW", "WHAT", "WHO", "WHOM", "WHOSE", "NOT", "NO", "YES", "TRUE", "FALSE", "NULL" };


            foreach (var word1 in words1)
            {
                foreach (var word2 in words2)
                {
                    if (string.Equals(word1, word2, StringComparison.OrdinalIgnoreCase)
                        && !commonWords.Contains(word1, StringComparer.OrdinalIgnoreCase)
                        && !ignoredWords.Contains(word1, StringComparer.OrdinalIgnoreCase)
                        )
                    {
                        commonWords.Add(word1);
                        //commonWordRowsInColumn1.Add($"Row {Array.IndexOf(words1, word1) + 1}");
                        //commonWordRowsInColumn2.Add($"Row {Array.IndexOf(words2, word2) + 1}");
                        //Console.WriteLine($"Common word found: {word1} in Row {Array.IndexOf(words1, word1) + 1} of Column 1 and Row {Array.IndexOf(words2, word2) + 1} of Column 2");
                    }
                }
            }

            // Create a new DataTable
            var table = new DataTable();

            // Add columns to the DataTable
            table.Columns.Add("Common Word");
            //table.Columns.Add("Row in Column 1");
            //table.Columns.Add("Row in Column 2");

            // Loop through the common words and add each row to the DataTable
            for (int i = 0; i < commonWords.Count; i++)
            {
                table.Rows.Add(commonWords[i]);
            }

            return table;
        }
    }
}