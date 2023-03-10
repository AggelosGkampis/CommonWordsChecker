using OfficeOpenXml;
using System.Data;
using System.Linq;

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
            var worksheet = workbook.Worksheets[0];
            var ignoredWords = new HashSet<string> { "THE", "I", "A", "YOU", "HE", "SHE", "IT", "WE", "THEY", "AN", "AND", "OR",
                                         "BUT", "SO", "IN", "ON", "AT", "IS", "ARE", "WAS", "WERE", "BE", "BEEN", "AM", "DO", "DOES", "DID", "HAVE",
                                         "HAS", "HAD", "OF", "TO", "FOR", "FROM", "BY", "WITH", "THAT", "THIS", "THOSE", "THESE", "IF", "THEN", "ELSE",
                                         "WHEN", "WHERE", "WHILE", "HOW", "WHAT", "WHO", "WHOM", "WHOSE", "NOT", "NO", "YES", "TRUE", "FALSE", "NULL", "foreign","Movie", "Film" };

            Dictionary<string, List<string>> column1Words;
            Dictionary<string, List<string>> column2Words;

            DictionariesWithSplitWords(ignoredWords, worksheet, out column1Words, out column2Words);

            DataTable table = CreateDatatableOutput(column1Words, column2Words, worksheet);

            return table;
        }



        public static void DictionariesWithSplitWords(HashSet<string> ignoredWords, ExcelWorksheet worksheet, out Dictionary<string, List<string>> column1Words, out Dictionary<string, List<string>> column2Words)
        {
            column1Words = new Dictionary<string, List<string>>();
            column2Words = new Dictionary<string, List<string>>();

            // Split words in columns 1 and 2 and track the row numbers where each word occurs
            for (int i = 1; i <= worksheet.Dimension.Rows; i++)
            {
                var column1Value = worksheet.Cells[i, 1].Value?.ToString()?.Trim()?.ToLower();
                var column2Value = worksheet.Cells[i, 2].Value?.ToString()?.Trim()?.ToLower();

                if (!string.IsNullOrEmpty(column1Value))
                {
                    var cellAddress = $"A{i}";
                    var words = column1Value.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries)
                                     .SelectMany(line => line.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries))
                                     .Where(w => !ignoredWords.Contains(w, StringComparer.OrdinalIgnoreCase))
                                     .ToList();
                   
                        column1Words[cellAddress] = words;
                    
                    
                }

                if (!string.IsNullOrEmpty(column2Value))
                {
                    var cellAddress = $"B{i}";
                    var words = column2Value.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries)
                                     .SelectMany(line => line.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries))
                                     .Where(w => !ignoredWords.Contains(w, StringComparer.OrdinalIgnoreCase))
                                     .ToList();
                 
                        column2Words[cellAddress] = words;
                    
                    
                }
            }
        }


        public static DataTable CreateDatatableOutput(Dictionary<string, List<string>> column1Words, Dictionary<string, List<string>> column2Words, ExcelWorksheet worksheet)
        {
            DataTable result = new DataTable();
            result.Columns.Add("Common Word");
            result.Columns.Add("Column 1 Matches");
            result.Columns.Add("Column 2 Matches");
            result.Columns.Add("Column 1 Values");
            result.Columns.Add("Column 2 Values");

            // Find common words between the two columns and track the row numbers where each word occurs
            foreach (var word in column1Words.SelectMany(p => p.Value).Intersect(column2Words.SelectMany(p => p.Value)).Distinct())
            {
                var column1Matches = new List<string>();
                var column2Matches = new List<string>();
                var column1Values = new HashSet<string>();  // use HashSet to eliminate duplicates
                var column2Values = new HashSet<string>();  // use HashSet to eliminate duplicates

                // Get the rows where the common word occurs in column 1
                foreach (var kvp in column1Words)
                {
                    if (kvp.Value.Contains(word))
                    {
                        var rowNumber = GetRowNumberFromCellAddress(kvp.Key);
                        column1Matches.Add(rowNumber);
                        if (!column1Values.Contains(word))
                        {
                            column1Values.Add(worksheet.Cells[kvp.Key].Value?.ToString() ?? "null");
                        }
                    }
                }

                // Get the rows where the common word occurs in column 2
                foreach (var kvp in column2Words)
                {
                    if (kvp.Value.Contains(word))
                    {
                        var rowNumber = GetRowNumberFromCellAddress(kvp.Key);
                        column2Matches.Add(rowNumber);
                        if (!column2Values.Contains(word))
                        {
                            column2Values.Add(worksheet.Cells[kvp.Key].Value?.ToString() ?? "null");
                        }
                    }
                }

                // Add a row for each occurrence of the common word
                for (int i = 0; i < Math.Max(column1Matches.Count, column2Matches.Count); i++)
                {
                    string column1Value = "";
                    string column2Value = "";

                    // Get the value from column 1 if it exists and has not been processed already
                    if (i < column1Matches.Count)
                    {
                        var rowNumber = column1Matches[i];
                        if (!column1Values.Contains(rowNumber))
                        {
                            column1Values.Add(rowNumber);
                            column1Value = worksheet.Cells[$"A{rowNumber}"].Value?.ToString() ?? "null";
                        }
                    }

                    // Get the value from column 2 if it exists and has not been processed already
                    if (i < column2Matches.Count)
                    {
                        var rowNumber = column2Matches[i];
                        if (!column2Values.Contains(rowNumber))
                        {
                            column2Values.Add(rowNumber);
                            column2Value = worksheet.Cells[$"B{rowNumber}"].Value?.ToString() ?? "null";
                        }
                    }

                    // Add the row to the output table
                    result.Rows.Add(word, string.Join(",", column1Matches), string.Join(",", column2Matches), column1Value, column2Value);
                }
            }

            return result;
        }






        private static string GetRowNumberFromCellAddress(string cellAddress)
        {
            return new string(cellAddress.Where(c => char.IsDigit(c)).ToArray());
        }

    }

}

