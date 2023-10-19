using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;

namespace docxParserForms.DocxHandler
{
    public static class DescriptionHandler
    {
        private const string _keyWords = "Картинка Рисунок Рис. Фигура Фиг. Изображение Image Figure";

        public static string? GetDescription(Paragraph paragraph)
        {
            var stringBuilder = new StringBuilder();
            foreach (var run in paragraph.Elements<Run>())
                stringBuilder.Append(run.InnerText);

            var splittedText = stringBuilder.ToString().Split(Environment.NewLine);
            foreach (var line in splittedText)
            {
                var splittedLine = line.Split(' ');
                if (splittedLine[0].Trim().Length > 3 && _keyWords.Contains(splittedLine[0]))
                    return TakeDataFromString(splittedLine, splittedLine[0]);
            }

            return null;
        }

        public static string TakeDataFromString(string[] splittedLine, string keyWord)
        {
            var sb = new StringBuilder();
            var separators = ".,;:-+*/\\–";

            for (var i = 0; i < splittedLine.Length; i++)
            {
                if (splittedLine[i].StartsWith(keyWord, StringComparison.OrdinalIgnoreCase)
                    || splittedLine[i].Length == 0
                    || int.TryParse(splittedLine[i], out _)
                    || separators.Contains(splittedLine[i].ToCharArray()[0].ToString()))
                    continue;
                else
                {
                    sb.Append(splittedLine[i]);
                    sb.Append(' ');
                }
            }

            if (sb.ToString().StartsWith("SEQ ARABIC"))
                sb.Remove(0, 10);

            return sb.ToString().Trim();
        }

        public static void HandleDescriptionToImages(List<string> descriptions, string _splitExample)
        {
            List<string> temp = new();
            temp.AddRange(descriptions);
            descriptions.Clear();

            var separators = _splitExample.Split("; ");
            separators = separators.Where(sep => sep != null || sep?.Length != 0).ToArray();

            int sepIndex = 0, curIndex = 0;

            for (int i = 0; i < temp.Count; i++)
            {
                if (temp[i].Contains(separators[sepIndex]))
                {
                    curIndex = temp[i].IndexOf(separators[sepIndex]);
                    while (true)
                    {
                        sepIndex++;

                        StringBuilder tempString = new();
                        var splittedDescr = temp[i].Split(' ');

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
                    descriptions.Add(temp[i]);
                }
            }
        }
    }
}
