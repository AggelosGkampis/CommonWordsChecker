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
            var commonWordRowsInColumn1 = new List<string>();
            var commonWordRowsInColumn2 = new List<string>();
            var worksheet = workbook.Worksheets[0];

            var column1 = string.Join(" ", worksheet.Cells["A:A"].Select(c => c.Value?.ToString().Trim().ToLower()));
            var column2 = string.Join(" ", worksheet.Cells["B:B"].Select(c => c.Value?.ToString().Trim().ToLower()));

            var words1 = column1.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            var words2 = column2.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var word1 in words1)
            {
                foreach (var word2 in words2)
                {
                    if (string.Equals(word1, word2, StringComparison.OrdinalIgnoreCase)
                        && !commonWords.Contains(word1))
                    {
                        commonWords.Add(word1);
                        commonWordRowsInColumn1.Add($"Row {Array.IndexOf(words1, word1) + 1}");
                        commonWordRowsInColumn2.Add($"Row {Array.IndexOf(words2, word2) + 1}");
                        Console.WriteLine($"Common word found: {word1} in Row {Array.IndexOf(words1, word1) + 1} of Column 1 and Row {Array.IndexOf(words2, word2) + 1} of Column 2");
                    }
                }
            }

            // Create a new DataTable
            var table = new DataTable();

            // Add columns to the DataTable
            table.Columns.Add("Common Word");
            table.Columns.Add("Row in Column 1");
            table.Columns.Add("Row in Column 2");

            // Loop through the common words and add each row to the DataTable
            for (int i = 0; i < commonWords.Count; i++)
            {
                table.Rows.Add(commonWords[i], commonWordRowsInColumn1[i], commonWordRowsInColumn2[i]);
            }

            return table;
        }

    }
}