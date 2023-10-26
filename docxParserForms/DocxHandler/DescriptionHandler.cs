using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;

namespace docxParserForms.DocxHandler
{
    public static class DescriptionHandler
    {
        private const string KeyWords = "Картинка Рисунок Рис. Фигура Фиг. Изображение Image Figure Fig. Picture Pic.";

        public static string? GetDescription(Paragraph paragraph)
        {
            var stringBuilder = new StringBuilder();
            foreach (var run in paragraph.Elements<Run>())
                stringBuilder.Append(run.InnerText);

            var splittedText = stringBuilder.ToString().Split(Environment.NewLine);
            foreach (var line in splittedText)
            {
                var splittedLine = line.Split(' ');
                if (splittedLine[0].Trim().Length > 3 && KeyWords.Contains(splittedLine[0]))
                    return TakeDataFromString(splittedLine, splittedLine[0]);
            }

            return null;
        }

        public static string TakeDataFromString(string[] splittedLine, string keyWord)
        {
            var sb = new StringBuilder();
            const string separators = ".,;:-+*/\\–";

            foreach (var line in splittedLine)
            {
                if (line.StartsWith(keyWord, StringComparison.OrdinalIgnoreCase)
                    || line.Length == 0
                    || double.TryParse(line, out _)
                    || separators.Contains(line.ToCharArray()[0].ToString()))
                    continue;
                else
                {
                    sb.Append(line);
                    sb.Append(' ');
                }
            }

            return sb.ToString().Replace("SEQ ARABIC", "").Trim();
        }

        public static void HandleDescriptionToImages(List<string> descriptions, string _splitExample)
        {
            List<string> temp = new();
            temp.AddRange(descriptions);
            descriptions.Clear();

            var separators = _splitExample.Split("; ");
            separators = separators.Where(sep => sep != null || sep?.Length != 0).ToArray();

            int sepIndex = 0, curIndex = 0;

            foreach (var t in temp)
            {
                if (t.Contains(separators[sepIndex]))
                {
                    curIndex = t.IndexOf(separators[sepIndex], StringComparison.Ordinal);
                    while (true)
                    {
                        sepIndex++;

                        StringBuilder tempString = new();
                        var splittedDescr = t.Split(' ');

                        while (curIndex < splittedDescr.Length && splittedDescr[curIndex] != separators[sepIndex])
                        {
                            tempString.Append(splittedDescr[curIndex++]).Append(' ');
                        }

                        descriptions.Add(DescriptionHandler.TakeDataFromString(tempString.ToString().Split(' '), separators[sepIndex]));
                        curIndex++;

                        if (curIndex >= splittedDescr.Length - 1)
                            break;
                    }
                }
                else
                {
                    sepIndex = 0;
                    descriptions.Add(t);
                }
            }
        }
    }
}