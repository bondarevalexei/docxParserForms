using System.Collections;
using System.Globalization;
using System.Text;
using static System.Double;

namespace docxParserForms.DocxHandler
{
    public static class DescriptionHandler
    {
        private static readonly string[] KeyWords =
            { "картинка", "рисунок", "рис.", "рис", "фигура", "фиг.", "фиг",
            "изображение", "image", "figure", "fig.", "fig", "picture", "pic.", "pic" };

        public static string? GetDescription(string line, Hashtable? descriptionsHash)
        {
            IFormatProvider formatter = new NumberFormatInfo { NumberDecimalSeparator = "." };
            var style = NumberStyles.Number | NumberStyles.AllowCurrencySymbol;

            var splittedLine = line.Split(' ');
            var anotherKeyWord = false;

            var splittedLine0 = splittedLine[0].Trim().Split('.');
            splittedLine0 = splittedLine0.Where(sr => true).ToArray();

            if (splittedLine0 != null && splittedLine0.Length >= 2
                && KeyWords.Contains(splittedLine0[0].ToLower())
                && (CheckForNumber(splittedLine0[1], style, formatter)
                || splittedLine0[1].Length >= 1
                && TryParse(splittedLine0[1], style, formatter, out _)))
                anotherKeyWord = true;


            if ((splittedLine[0].Trim().Length <= 3 || !KeyWords.Contains(splittedLine[0].ToLower())) && !anotherKeyWord)
                return null;

            var keyWord = !anotherKeyWord
                ? splittedLine[0] + " " + splittedLine[1]
                : splittedLine0[0] + ". " + splittedLine0[1];

            keyWord = keyWord.ToLower();

            var keyForFunc = !anotherKeyWord
                ? splittedLine[0] : splittedLine0[0];

            if (descriptionsHash == null || !descriptionsHash.ContainsKey(keyWord))
                return TakeDataFromString(splittedLine, keyForFunc, style, formatter);

            return descriptionsHash[keyWord]?.ToString();
        }

        public static bool IsOnlyKey(string line)
        {
            var splittedLine = line.Trim().Split(' ');
            return splittedLine.Length <= 3
                && splittedLine[0].Trim().Length >= 3
                || KeyWords.Contains(splittedLine[0].ToLower());
        }

        public static string TakeDataFromString(string[] splittedLine, string keyWord, NumberStyles style, IFormatProvider formatter)
        {
            var sb = new StringBuilder();
            const string separators = ".,;:-+*/\\–";

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

        public static bool CheckForNumber(string line, NumberStyles style, IFormatProvider formatter)
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

            IFormatProvider formatter = new NumberFormatInfo { NumberDecimalSeparator = "." };
            var style = NumberStyles.Number | NumberStyles.AllowCurrencySymbol;

            var separators = _splitExample.Split("; ");
            separators = separators.Where(sep => sep != null || sep?.Length != 0).ToArray();

            int sepIndex = 0, curIndex = 0;

            if (description.Contains(separators[sepIndex]))
            {
                var splittedDescr = description.Split(' ');

                foreach (var item in splittedDescr)
                {
                    if (string.Compare(item, separators[sepIndex]) == 0) break;
                    curIndex++;
                }

                curIndex = description.IndexOf(separators[sepIndex], StringComparison.Ordinal);
                while (true)
                {
                    sepIndex++;
                    StringBuilder tempString = new();

                    while (curIndex < splittedDescr.Length && splittedDescr[curIndex] != separators[sepIndex])
                    {
                        tempString.Append(splittedDescr[curIndex++]).Append(' ');
                    }

                    res.Add(DescriptionHandler.TakeDataFromString(
                        tempString.ToString().Split(' '), separators[sepIndex], style, formatter));
                    curIndex++;

                    if (curIndex >= splittedDescr.Length - 1)
                        break;
                }
            }

            return res;
        }
    }
}