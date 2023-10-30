using System.Collections;
using System.Globalization;
using System.Text;
using SautinSoft.Document;
using static System.String;

namespace docxParserForms.DocxHandler
{
    public static class TextHandler
    {
        private static readonly string[] KeyWords =
            { "картинка", "рисунок", "рис.", "фигура", "фиг.", "изображение", "image", "figure", "fig.", "picture", "pic." };
        private static readonly string[] KeyWordsAnotherCase =
            { "картинке", "рисунке", "рис.", "фигуре", "фиг.", "изображении", "image", "figure", "fig.", "picture", "pic." };

        public static Hashtable CollectDescriptionsFromText(string filepath)
        {
            var dc = DocumentCore.Load(filepath);
            Hashtable descriptionsHashtable = new();

            foreach (var run in dc.GetChildElements(true, ElementType.Run))
            {
                try
                {
                    if (Compare(run.Content.ToString(), "- рис.") == 0)
                    {
                        Console.Write("pora");
                    }

                    var splittedRun = run.Content.ToString().Split(' ');
                    splittedRun = splittedRun.Where((sr) => Compare(sr, "") != 0 || sr != null).ToArray();

                    var isAnotherCase = 0;
                    if (CheckFirstWord(splittedRun)) isAnotherCase = 1;
                    else if (CheckFirstWords(splittedRun)) isAnotherCase = 2;

                    if (isAnotherCase == 0) continue;

                    var (key, value) = GetDescrFromLine(splittedRun, isAnotherCase);
                    if (value.Length != 0 && !descriptionsHashtable.Contains(key))
                        descriptionsHashtable.Add(key, value);
                }
                catch (Exception ex)
                {
                    throw new Exception(ex.Message);
                }
            }

            return descriptionsHashtable;
        }

        private static bool CheckFirstWord(IReadOnlyList<string> splittedRun)
        {
            if (splittedRun.Count < 3) return false;

            if (splittedRun[0].Length > 3 && KeyWords.Contains(splittedRun[0].ToLower())) return true;
            return CompareOrdinal(splittedRun[0], "-") == 0
                   && splittedRun[1].Length > 3 && KeyWords.Contains(splittedRun[1].ToLower());
        }

        private static bool CheckFirstWords(IReadOnlyList<string> splittedRun)
        {
            if (splittedRun.Count < 4) return false;

            var index = 0;
            if (CompareOrdinal(splittedRun[0], "-") == 0) index = 1;

            return (splittedRun.Count >= index + 2 && Compare(splittedRun[index].ToLower(), "на",
                                               StringComparison.OrdinalIgnoreCase) == 0
                                           && KeyWordsAnotherCase.Contains(splittedRun[index + 1].ToLower())) ||
               (splittedRun.Count >= index + 3 && Compare(splittedRun[index].ToLower(), "on",
                                           StringComparison.OrdinalIgnoreCase) == 0
                                       && Compare(splittedRun[index + 1].ToLower(), "the",
                                           StringComparison.OrdinalIgnoreCase) == 0
                                       && KeyWordsAnotherCase.Contains(splittedRun[index + 2].ToLower()));
        }

        private static (string, string) GetDescrFromLine(IReadOnlyList<string> splittedRun, int isAnotherCase)
        {
            var sb = new StringBuilder();
            const string separators = ".,;:-+*/\\–";

            IFormatProvider formatter = new NumberFormatInfo { NumberDecimalSeparator = "." };
            const NumberStyles style = NumberStyles.Number | NumberStyles.AllowCurrencySymbol;
            for (var i = isAnotherCase; i < splittedRun.Count; i++)
            {
                if (splittedRun[i].Length == 0
                    || (double.TryParse(splittedRun[i], style, formatter, out _) || CheckForNumber(splittedRun[i], style, formatter)) && i == isAnotherCase
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
            if (line.Length >= 1) return line.EndsWith('.') && double.TryParse(line[..^1], style, formatter, out _);
            return false;
        }
    }
}
