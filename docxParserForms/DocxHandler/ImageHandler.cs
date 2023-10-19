using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace docxParserForms.DocxHandler
{
    public static class ImageHandler
    {
        public static Bitmap ExtractImage(MainDocumentPart wDoc, Drawing image, List<string> imageTypes)
        {
            var imageFirst = image.Inline.Graphic.GraphicData
                .Descendants<DocumentFormat.OpenXml.Drawing.Pictures
                    .Picture>().FirstOrDefault();

            var blip = imageFirst.BlipFill.Blip.Embed.Value;

            ImagePart img = (ImagePart)wDoc.Document.MainDocumentPart
                .GetPartById(blip);
            imageTypes.Add(img.ContentType);

            using Image resultImage = Bitmap.FromStream(img.GetStream());
            return new Bitmap(resultImage);
        }
    }
}
