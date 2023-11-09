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

        private static readonly IFormatProvider Formatter = new NumberFormatInfo { NumberDecimalSeparator = "." };
        private const NumberStyles Style = NumberStyles.Number | NumberStyles.AllowCurrencySymbol;

        public static SortedDictionary<string, string> CollectDescriptionsFromText(string filepath)
        {
            var dc = DocumentCore.Load(filepath);
            SortedDictionary<string,string> descriptionsDict = new();

            foreach (var element in dc.GetChildElements(true, ElementType.Run))
            {
                var run = (Run)element;

                try
                {
                    var splittedRun = run.Text.Split(' ');
                    splittedRun = splittedRun.Where(sr => true).ToArray();

                    var isAnotherCase = 0;
                    if (CheckFirstWordAndNumber(splittedRun)) isAnotherCase = 1;
                    else if (CheckFirstWordsAndNumber(splittedRun)) isAnotherCase = 2;

                    if (isAnotherCase == 0) continue;

                    var (key, value) = GetDescrFromLine(splittedRun, isAnotherCase);
                    if (value.Length != 0 && !descriptionsDict.ContainsKey(key))
                        descriptionsDict.Add(key, value);
                }
                catch (Exception ex)
                {
                    throw new Exception(ex.Message);
                }
            }

            return descriptionsDict;
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
                    || separators.Contains(splittedRun[i].ToCharArray()[0].ToString()))
                    continue;

                sb.Append(splittedRun[i]);
                sb.Append(' ');
            }

            var isDashFirst = CompareOrdinal(splittedRun[0], "-") == 0;

            return !isDashFirst
                ? (splittedRun[0] + " " + splittedRun[1], sb.ToString().Replace("SEQ ARABIC", "").Trim())
                : (splittedRun[1] + " " + splittedRun[2], sb.ToString().Replace("SEQ ARABIC", "").Trim());
        }

        private static bool CheckForNumber(string line, NumberStyles style, IFormatProvider formatter)
        {
            if (line.Length >= 1) return line.EndsWith('.') && TryParse(line[..^1], style, formatter, out _) || TryParse(line, Style, Formatter, out _);
            return false;
        }
    }
}
