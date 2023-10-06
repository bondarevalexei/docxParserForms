using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Drawing.Imaging;
using System.Diagnostics;

namespace docxParserForms.DocxHandler
{
    public class MainHandler
    {
        public string ReadText(string filepath)
        {
            string text = String.Empty;

            using (WordprocessingDocument wordDocument =
                WordprocessingDocument.Open(filepath, false))
            {
                Body body = wordDocument.MainDocumentPart.Document.Body;
                foreach (var paragraph in body.Elements<Paragraph>())
                {
                    foreach (var run in paragraph.Elements<Run>())
                    {
                        text += run.InnerText;
                    }
                    text += Environment.NewLine;
                }

                var splittedText = text.Split(Environment.NewLine);

                foreach (var line in splittedText)
                {
                    if (line.StartsWith("Рисунок"))
                        ExtractImages(body, wordDocument.MainDocumentPart);
                }
            }

            return text;
        }

        private List<string> ExtractImages(Body content, MainDocumentPart wDoc)
        {
            List<string> imageList = new();

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

                        string folder = Path.GetDirectoryName(
                            Process.GetCurrentProcess().MainModule.FileName);

                        ImagePart img = (ImagePart)wDoc.Document.MainDocumentPart
                            .GetPartById(blip);

                        string imageFileName = string.Empty;

                        using (Image toSaveImage = Bitmap.FromStream(img.GetStream()))
                        {
                            imageFileName = folder + @"\TestFiles\TestExtractor_" 
                                + DateTime.Now.Month.ToString().Trim() 
                                    + DateTime.Now.Day.ToString() 
                                        + DateTime.Now.Year.ToString() 
                                            + DateTime.Now.Hour.ToString() 
                                                + DateTime.Now.Minute.ToString()
                                                    + DateTime.Now.Second.ToString() 
                                                        + DateTime.Now.Millisecond.ToString() 
                                                            + ".png";

                            try
                            {
                                toSaveImage.Save(imageFileName, ImageFormat.Png);
                            }
                            catch (Exception ex) 
                            {
                                MessageBox.Show(ex.Message);
                            }
                        }

                        imageList.Add(imageFileName);
                    }

                }
            }

            return imageList;
        }
    }
}
