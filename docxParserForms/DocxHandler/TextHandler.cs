using System.Collections;
using System.Globalization;
using System.Text;
using SautinSoft.Document;
using static System.Double;
using static System.String;

namespace docxParserForms.DocxHandler
{
    public static class TextHandler
    {
        private static readonly string[] KeyWords =
            { "картинка", "рисунок", "рис.", "фигура", "фиг.", "изображение", "image", "figure", "fig.", "picture", "pic." };
        private static readonly string[] KeyWordsAnotherCase =
            { "картинке", "рисунке", "рис.", "фигуре", "фиг.", "изображении", "image", "figure", "fig.", "picture", "pic." };
        private static readonly string[] TriggerWords = { "изображен", "изображены" };

        private static readonly IFormatProvider Formatter = new NumberFormatInfo { NumberDecimalSeparator = "." };
        private const NumberStyles Style = NumberStyles.Number | NumberStyles.AllowCurrencySymbol;

        public static Hashtable CollectDescriptionsFromText(string filepath)
        {
            var dc = DocumentCore.Load(filepath);
            Hashtable descriptionsHash = new();

            foreach (var el in dc.GetChildElements(true, ElementType.Paragraph))
            {
                var paragraph = (Paragraph)el;
                var childElements = paragraph.GetChildElements(true).ToList();

                CheckRunsForDescriptions(childElements, descriptionsHash);
            }

            return descriptionsHash;
        }

        private static void CheckRunsForDescriptions(List<Element> childElements, Hashtable descriptionsHash)
        {
            bool isKeyUsed = false;
            StringBuilder sb = new();
            var tempIndex = 0;

            foreach (var child in childElements)
            {
                if (child.ElementType == ElementType.Run)
                {
                    var run = (Run)child;

                    if (IsOnlyKey(run.Text) || tempIndex < 3 && isKeyUsed)
                    {
                        isKeyUsed = true;
                        sb.Append(run.Text.Trim());
                        sb.Append(' ');
                        tempIndex++;
                        continue;
                    }

                    string? description = !isKeyUsed
                        ? run.Text : sb.ToString().Trim();

                    if (description != null && TryToAddDescription(description, descriptionsHash))
                        continue;

                    isKeyUsed = false;
                    tempIndex = 0;
                    sb.Clear();
                }
            }

            if (isKeyUsed && sb.Length != 0)
            {
                string? description = sb.ToString().Trim();

                if (description != null && TryToAddDescription(description, descriptionsHash))
                    return;
            }
        }

        private static bool TryToAddDescription(string description, Hashtable descriptionsHash)
        {
            try
            {
                var splittedRun = description.Trim().Split(' ');
                splittedRun = splittedRun.Where(sr => true).ToArray();

                var isAnotherCase = 0;
                if (CheckFirstWordAndNumber(splittedRun)) isAnotherCase = 1;
                else if (CheckFirstWordsAndNumber(splittedRun)) isAnotherCase = 2;

                if (isAnotherCase == 0) return true;

                var (key, value) = GetDescrFromLine(splittedRun, isAnotherCase);
                if (value.Length != 0 && !descriptionsHash.ContainsKey(key))
                    descriptionsHash.Add(key, value);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }

            return false;
        }

        private static bool IsOnlyKey(string text)
        {
            var splittedLine = text.Trim().Split(' ');

            if(splittedLine.Length == 0 || splittedLine.Length == 1) return false;

            if (Compare(splittedLine[0].ToLower(), "на") == 0 || Compare(splittedLine[0].ToLower(), "on") == 0)
            {
                return splittedLine.Length <= 3 && splittedLine[1].Trim().Length >= 3 || KeyWords.Contains(splittedLine[1].ToLower());
            }
            
            return splittedLine.Length <= 3 && splittedLine[0].Trim().Length >= 3 || KeyWords.Contains(splittedLine[0].ToLower());
        }

        private static bool CheckFirstWordAndNumber(IReadOnlyList<string> splittedRun)
        {
            if (splittedRun.Count < 3) return false;

            if (splittedRun[0].Length > 3 && KeyWords.Contains(splittedRun[0].ToLower())
                                          && CheckForNumber(splittedRun[1], Style, Formatter)) return true;
            return CompareOrdinal(splittedRun[0], "-") == 0
                   && splittedRun[1].Length > 3 && KeyWords.Contains(splittedRun[1].ToLower())
                   && CheckForNumber(splittedRun[2], Style, Formatter);
        }

        private static bool CheckFirstWordsAndNumber(IReadOnlyList<string> splittedRun)
        {
            if (splittedRun.Count < 4) return false;

            var index = 0;
            if (CompareOrdinal(splittedRun[0], "-") == 0) index = 1;

            return splittedRun.Count >= index + 2 && Compare(splittedRun[index].ToLower(), "на",
                                                      StringComparison.OrdinalIgnoreCase) == 0
                                                  && KeyWordsAnotherCase.Contains(splittedRun[index + 1].ToLower())
                                                  && CheckForNumber(splittedRun[index + 2], Style, Formatter) ||
               splittedRun.Count >= index + 3 && Compare(splittedRun[index].ToLower(), "on",
                                           StringComparison.OrdinalIgnoreCase) == 0
                                       && Compare(splittedRun[index + 1].ToLower(), "the",
                                           StringComparison.OrdinalIgnoreCase) == 0
                                       && KeyWordsAnotherCase.Contains(splittedRun[index + 2].ToLower())
                                       && CheckForNumber(splittedRun[index + 3], Style, Formatter);
        }

        private static (string, string) GetDescrFromLine(IReadOnlyList<string> splittedRun, int isAnotherCase)
        {
            var sb = new StringBuilder();
            const string separators = ".,;:-+*/\\–";

            for (var i = isAnotherCase; i < splittedRun.Count; i++)
            {
                if (splittedRun[i].Length == 0
                    || (TryParse(splittedRun[i], Style, Formatter, out _) || CheckForNumber(splittedRun[i], Style, Formatter)) && i == isAnotherCase
                    || separators.Contains(splittedRun[i].ToCharArray()[0].ToString())
                    || TriggerWords.Contains(splittedRun[i].ToLower()))
                    continue;

                sb.Append(splittedRun[i]);
                sb.Append(' ');
            }

            if(CompareOrdinal(splittedRun[0], "-") == 0)
                _ = splittedRun[0].Replace('-', ' ');

            splittedRun = splittedRun.Where(sr => true).ToArray();

            return (splittedRun[isAnotherCase - 1] + " " + splittedRun[isAnotherCase], sb.ToString().Replace("SEQ ARABIC", "").Trim());
        }

        private static bool CheckForNumber(string line, NumberStyles style, IFormatProvider formatter)
        {
            if (line.Length >= 1) return line.EndsWith('.') && TryParse(line[..^1], style, formatter, out _) || TryParse(line, Style, Formatter, out _);
            return false;
        }
    }
}
