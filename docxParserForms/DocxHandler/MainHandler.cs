using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;
using System.Data.SqlClient;
using System.Drawing.Imaging;

namespace docxParserForms.DocxHandler
{
    public class MainHandler
    {
        private readonly string _connectionString;

        public MainHandler(string connectionString) =>
            _connectionString = connectionString;

        public void HandleFile(string filepath)
        {
            List<Bitmap> images = new();
            List<string> descriptions = new();

            try
            {
                using (WordprocessingDocument wordDocument =
                    WordprocessingDocument.Open(filepath, false))
                {
                    Body body = wordDocument.MainDocumentPart.Document.Body;
                    bool flag = false;

                    Bitmap? imageBitmap = null;
                    string? description = null;
                    int paragraphCounter = 0, imageCounter = 0;

                    foreach (Paragraph paragraph in body.Descendants<Paragraph>())
                    {
                        paragraphCounter++;
                        ParagraphProperties paragraphProperties = paragraph.ParagraphProperties;
                        foreach (Run run in paragraph.Descendants<Run>())
                        {
                            Drawing image =
                                run.Descendants<Drawing>().FirstOrDefault();

                            if (image != null && image.Inline != null)
                            {
                                imageBitmap = new Bitmap(ExtractImage(wordDocument.MainDocumentPart, image));
                            }
                            else
                            {
                                flag = true;
                                break;
                            }
                        }

                        if (flag)
                        {
                            description = GetDescription(paragraph);
                            flag = false;
                        }

                        if (imageBitmap != null && description != null)
                        {
                            images.Add(imageBitmap);
                            descriptions.Add(description);

                            (imageBitmap, description) = (null, null);
                        }

                        if(paragraphCounter > 1 && flag)
                            (imageBitmap, description) = (null, null);
                    }
                }

                SaveToDb(descriptions, images);
                MessageBox.Show($"Файл {filepath} успешно обработан. Добавлено {descriptions.Count} элемента(ов).");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private string? GetDescription(Paragraph paragraph)
        {
            var stringBuilder = new StringBuilder();

            foreach (var run in paragraph.Elements<Run>())
                stringBuilder.Append(run.InnerText);

            var splittedText = stringBuilder.ToString().Split(Environment.NewLine);

            foreach (var line in splittedText)
                if (line.StartsWith("Рисунок", StringComparison.OrdinalIgnoreCase))
                    return TakeDataFromString(line);

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



        private static string TakeDataFromString(string line)
        {
            var splittedLine = line.Split(' ');
            bool flag = false;
            var sb = new StringBuilder();

            for (var i = 0; i < splittedLine.Length; i++)
            {
                if (int.TryParse(splittedLine[i], out _))
                {
                    flag = true;
                    continue;
                }

                if (flag)
                    sb.Append(splittedLine[i] + " ");
            }

            if (!flag) return line[9..].Trim();
            return sb.ToString();
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