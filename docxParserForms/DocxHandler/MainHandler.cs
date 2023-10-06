﻿using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Drawing.Imaging;
using System.Diagnostics;
using docxParserForms.Db.Context;

namespace docxParserForms.DocxHandler
{
    public class MainHandler
    {
        private ModelsDbContext _db = new ModelsDbContext();

        public string ReadText(string filepath)
        {
            string text = string.Empty;

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

                List<string> images = new ();
                int imageCounter = 0;
                foreach (var line in splittedText)
                {
                    if (line.StartsWith("Рисунок"))
                        imageCounter++;
                }

                images = ExtractImages(body, wordDocument.MainDocumentPart);
                foreach(var image in images)
                {
                    text += image + Environment.NewLine;
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
