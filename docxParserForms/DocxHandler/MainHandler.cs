using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace docxParserForms.DocxHandler
{
    public class MainHandler
    {
        public string ReadText(string filepath)
        {
            string text = String.Empty;

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filepath, false))
            {
                Body body = wordDocument.MainDocumentPart.Document.Body;
                foreach(var paragraph in body.Elements<Paragraph>())
                {
                    foreach (var run in paragraph.Elements<Run>())
                    {
                        text += run.InnerText;
                    }
                    text += Environment.NewLine;
                }
            }

            return text;
        }
    }
}
