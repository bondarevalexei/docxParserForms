using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;
using System.Data.SqlClient;
using System.Drawing.Imaging;
using Newtonsoft.Json.Linq;

namespace docxParserForms.DocxHandler
{
    public class MainHandler
    {
        private const string _keyWords = "Картинка Рисунок Рис. Фигура Фиг. Изображение Image Figure";
        private readonly string _connectionString;
        private readonly string _splitExample;

        public MainHandler()
        {
            using (StreamReader sr = new("./../../../appsettings.json"))
            {
                string json = sr.ReadToEnd();
                dynamic data = JObject.Parse(json);
                _connectionString = data.connectionString;
                _splitExample = data.splitExample;
            }
        }

        public List<Model> HandleFile(string filepath)
        {
            List<Bitmap> images = new();
            List<string> descriptions = new();
            List<Model> models = new();

            try
            {
                using (WordprocessingDocument wordDocument =
                    WordprocessingDocument.Open(filepath, false))
                {
                    HandleParagraphsInBody(wordDocument, images, descriptions);
                }

                WriteDataInModelsList(images, descriptions, models);

                //SaveToDb(descriptions, images);
                MessageBox.Show($"Файл {filepath} успешно обработан. Добавлено {descriptions.Count} элемента(ов).");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            return models;
        }

        private void WriteDataInModelsList(List<Bitmap> images,
            List<string> descriptions, List<Model> models)
        {
            for (int i = 0; i < descriptions.Count; i++)
                models.Add(new Model { Description = descriptions[i], Image = images[i] });
        }

        private void HandleParagraphsInBody(WordprocessingDocument wordDocument,
            List<Bitmap> images, List<string> descriptions)
        {
            Body body = wordDocument.MainDocumentPart.Document.Body;

            List<Bitmap> imagesInPargraph = new();
            List<string> descriptionsInParagraph = new();

            Bitmap? imageBitmap = null;
            string? description = null;
            int paragraphCounter = 0;

            foreach (Paragraph paragraph in body.Descendants<Paragraph>())
            {
                paragraphCounter++;

                foreach (Run run in paragraph.Descendants<Run>())
                {
                    Drawing image =
                        run.Descendants<Drawing>().FirstOrDefault();

                    if (image != null && image.Inline != null)
                    {
                        imageBitmap = new Bitmap(ExtractImage(wordDocument.MainDocumentPart, image));
                        imagesInPargraph.Add(imageBitmap);
                    }
                    else
                    {
                        description = GetDescription(paragraph);
                        if (description == null || description?.Trim().Length == 0)
                            continue;

                        descriptionsInParagraph.Add(description);
                        break;
                    }
                }

                descriptionsInParagraph = descriptionsInParagraph.Where(description => description != null).ToList();
                if (descriptionsInParagraph.Count == 0)
                    continue;

                if (descriptionsInParagraph.Count < imagesInPargraph.Count)
                    HandleDescriptionToImages(descriptionsInParagraph);

                if (descriptionsInParagraph.Count != imagesInPargraph.Count)
                {
                    (imageBitmap, description) = (null, null);
                    imagesInPargraph.Clear();
                    descriptionsInParagraph.Clear();
                }


                if (AddNewLinesToLists(images, descriptions, imagesInPargraph, descriptionsInParagraph)
                    || paragraphCounter > 1)
                {
                    (imageBitmap, description) = (null, null);
                    imagesInPargraph.Clear();
                    descriptionsInParagraph.Clear();
                }
            }
        }

        private void HandleDescriptionToImages(List<string> descriptions)
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

                        descriptions.Add(TakeDataFromString(tempString.ToString().Split(' '), separators[sepIndex]));
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

        private bool AddNewLinesToLists(List<Bitmap> images,
            List<string> descriptions, List<Bitmap> tempImages, List<string> tempDescriptions)
        {
            foreach (var description in tempDescriptions)
                descriptions.Add(description);

            foreach (var image in tempImages)
                images.Add(image);

            return true;
        }

        private string? GetDescription(Paragraph paragraph)
        {
            var stringBuilder = new StringBuilder();

            foreach (var run in paragraph.Elements<Run>())
                stringBuilder.Append(run.InnerText);

            var splittedText = stringBuilder.ToString().Split(Environment.NewLine);

            foreach (var line in splittedText)
            {
                var splittedLine = line.Split(' ');
                if (_keyWords.Contains(splittedLine[0]))
                    return TakeDataFromString(splittedLine, splittedLine[0]);
            }

               // if (line.StartsWith("Рисунок", StringComparison.OrdinalIgnoreCase))
                 //   return TakeDataFromString(line);

            return null;
        }

        private Bitmap ExtractImage(MainDocumentPart wDoc, Drawing image)
        {
            var imageFirst = image.Inline.Graphic.GraphicData
                .Descendants<DocumentFormat.OpenXml.Drawing.Pictures
                    .Picture>().FirstOrDefault();

            var blip = imageFirst.BlipFill.Blip.Embed.Value;

            ImagePart img = (ImagePart)wDoc.Document.MainDocumentPart
                .GetPartById(blip);

            using Image resultImage = Bitmap.FromStream(img.GetStream());
            return new Bitmap(resultImage);
        }

        private string TakeDataFromString(string[] splittedLine, string keyWord)
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

        private void SaveToDb(List<string> descriptions, List<Bitmap> images)
        {
            var count = descriptions.Count;

            for (int i = 0; i < count; i++)
            {
                using (var connection = new SqlConnection(_connectionString))
                {
                    using (var command = connection.CreateCommand())
                    {
                        command.CommandText =
                            "INSERT INTO parser_db (description, image) VALUES (@description, @image)";

                        using var memoryStream = new MemoryStream();
                        {
                            images[i].Save(memoryStream, ImageFormat.Jpeg);
                            memoryStream.Position = 0;

                            var sqlImageParam = new SqlParameter(
                                "@image", System.Data.SqlDbType.VarBinary, (int)memoryStream.Length)
                            {
                                Value = memoryStream.ToArray()
                            };

                            var sqlDescriptionParam = new SqlParameter(
                                "@description", System.Data.SqlDbType.VarChar, descriptions[i].Length)
                            {
                                Value = descriptions[i].ToString()
                            };

                            command.Parameters.Add(sqlDescriptionParam);
                            command.Parameters.Add(sqlImageParam);
                        }
                        connection.Open();
                        command.ExecuteNonQuery();
                    }
                }
            }
        }
    }
}