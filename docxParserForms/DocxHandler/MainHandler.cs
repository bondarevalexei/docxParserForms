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

                    descriptions = GetDescriptions(body);
                    images = ExtractImages(body, wordDocument.MainDocumentPart);
                }

                SaveToDb(descriptions, images);

                MessageBox.Show($"Файл {filepath} успешно обработан. Добавлено {descriptions.Count} элементов.");
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public List<string> GetDescriptions(Body body)
        {
            List<string> descriptions = new();
            string tempText = string.Empty;
            foreach (var paragraph in body.Elements<Paragraph>())
            {
                foreach (var run in paragraph.Elements<Run>())
                    tempText += run.InnerText;
                tempText += Environment.NewLine;
            }

            var splittedText = tempText.Split(Environment.NewLine);

            foreach (var line in splittedText)
                if (line.StartsWith("Рисунок"))
                    descriptions.Add(TakeDataFromString(line));

            return descriptions;
        }

        private List<Bitmap> ExtractImages(Body content, MainDocumentPart wDoc)
        {
            List<Bitmap> imageList = new();

            foreach (Paragraph par in content.Descendants<Paragraph>())
            {
                ParagraphProperties paragraphProperties = par.ParagraphProperties;
                foreach (Run run in par.Descendants<Run>())
                {
                    Drawing image =
                        run.Descendants<Drawing>().FirstOrDefault();

                    if (image != null)
                    {
                        var imageFirst = image.Inline.Graphic.GraphicData
                            .Descendants<DocumentFormat.OpenXml.Drawing.Pictures
                                .Picture>().FirstOrDefault();

                        var blip = imageFirst.BlipFill.Blip.Embed.Value;

                        ImagePart img = (ImagePart)wDoc.Document.MainDocumentPart
                            .GetPartById(blip);

                        using Image resultImage = Bitmap.FromStream(img.GetStream());
                        imageList.Add(new Bitmap(resultImage));
                    }
                }
            }

            return imageList;
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
                        command.CommandText = "INSERT INTO parser_db (description, image) VALUES (@description, @image)";

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