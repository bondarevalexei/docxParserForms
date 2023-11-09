﻿using System.Collections;
using System.Globalization;
using System.Text;
using SautinSoft.Document;
using static System.Double;
// using Run = DocumentFormat.OpenXml.Wordprocessing.Run;

namespace docxParserForms.DocxHandler
{
    public static class DescriptionHandler
    {
        private static readonly string[] KeyWords = { "картинка", "рисунок", "рис.", "фигура", "фиг.", "изображение", "image", "figure", "fig.", "picture", "pic." };

        public static string? GetDescription(Paragraph paragraph, Hashtable? descriptionsHash, ref bool isHashUsed)
        {
            var stringBuilder = new StringBuilder();
            foreach (var element in paragraph.GetChildElements(false, ElementType.Run))
            {
                var run = (Run)element;
                stringBuilder.Append(run.Text);
            }

            var splittedText = stringBuilder.ToString().Split(Environment.NewLine);
            foreach (var line in splittedText)
            {
                var splittedLine = line.Split(' ');
                if (splittedLine[0].Trim().Length <= 3 || !KeyWords.Contains(splittedLine[0].ToLower())) continue;

                var keyWord = splittedLine[0] + " " + splittedLine[1];
                if (descriptionsHash == null || !descriptionsHash.ContainsKey(keyWord))
                    return TakeDataFromString(splittedLine, splittedLine[0]);

                isHashUsed = true;
                return descriptionsHash[keyWord]?.ToString();

            }

            return null;
        }

        public static string TakeDataFromString(string[] splittedLine, string keyWord)
        {
            var sb = new StringBuilder();
            const string separators = ".,;:-+*/\\–";
            IFormatProvider formatter = new NumberFormatInfo { NumberDecimalSeparator = "." };
            var style = NumberStyles.Number | NumberStyles.AllowCurrencySymbol;

            foreach (var line in splittedLine)
            {
                if (line.StartsWith(keyWord, StringComparison.OrdinalIgnoreCase)
                    || CheckForNumber(line, style, formatter)
                    || line.Length == 0
                    || TryParse(line, style, formatter, out _)
                    || separators.Contains(line.ToCharArray()[0].ToString()))
                    continue;

                sb.Append(line);
                sb.Append(' ');
            }

            return sb.ToString().Replace("SEQ ARABIC", "").Trim();
        }

        private static bool CheckForNumber(string line, NumberStyles style, IFormatProvider formatter)
        {
            if (line.Length >= 1) return line.EndsWith('.') && TryParse(line[..^1], style, formatter, out _);
            return false;
        }

        public static bool IsDescriptionDevided(string description, string splitExample)
        {
            var separators = splitExample.Split("; ");
            separators = separators.Where(sep => sep != null || sep?.Length != 0).ToArray();

            if (description.Contains(separators[0])) return true;

            return false;
        }

        public static List<string> HandleDescriptionToMany(string description, string _splitExample)
        {
            List<string> res = new();

            var separators = _splitExample.Split("; ");
            separators = separators.Where(sep => sep != null || sep?.Length != 0).ToArray();

            int sepIndex = 0, curIndex = 0;

            if (description.Contains(separators[sepIndex]))
            {
                curIndex = description.IndexOf(separators[sepIndex], StringComparison.Ordinal);
                while (true)
                {
                    sepIndex++;

                    StringBuilder tempString = new();
                    var splittedDescr = description.Split(' ');

                    while (curIndex < splittedDescr.Length && splittedDescr[curIndex] != separators[sepIndex])
                    {
                        tempString.Append(splittedDescr[curIndex++]).Append(' ');
                    }

                    res.Add(DescriptionHandler.TakeDataFromString(tempString.ToString().Split(' '), separators[sepIndex]));
                    curIndex++;

                    if (curIndex >= splittedDescr.Length - 1)
                        break;
                }
            }

            return res;
        }
    }
}