using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml;

namespace ExcelCommonWords
{
    public class CommonWordsCheck
    {
        public static void Main(string[] args)
        {
            var commonWords = new List<string>();
            var commonWordRowsInColumn1 = new List<string>();
            var commonWordRowsInColumn2 = new List<string>();
            var file = new ExcelPackage(new System.IO.FileInfo(@"C:\Users\Isocratis-4\source\repos\CommonWordsChecker\CommonWordsChecker\test1.xlsx"));
            var worksheet = file.Workbook.Worksheets[0];

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


            file.Dispose();
        }
    }
}
