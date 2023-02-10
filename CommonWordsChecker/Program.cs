using System;
using System.Collections.Generic;

namespace ExcelCommonWords
{
    class Program
    {
        static void Main(string[] args)
        {
            string[] column1 = new string[] { "" };
            string[] column2 = new string[] { "δσα omega", "α " };
            List<string> commonWords = new List<string>();

            for (int i = 0; i < column1.Length; i++)
            {
                string[] words1 = column1[i].Split(' ');
                foreach (string word1 in words1)
                {
                    for (int j = 0; j < column2.Length; j++)
                    {
                        string[] words2 = column2[j].Split(' ');
                        foreach (string word2 in words2)
                        {
                            if (word1.ToLower() == word2.ToLower() && !commonWords.Contains(word1.ToLower()))
                            {
                                commonWords.Add(word1.ToLower());
                                Console.WriteLine($"Common word found: {word1}");
                            }
                        }
                    }
                }
            }

        }

    }
}


