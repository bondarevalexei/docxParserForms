using SautinSoft.Document;
using SautinSoft.Document.Drawing;

namespace docxParserForms.DocxHandler
{
    public static class ImageHandler
    {
        public static void ExtractImages(string path, List<Bitmap> result, List<string> imageTypes)
        {
            List<ImageData> imgInventory = new List<ImageData>();

            DocumentCore dc = DocumentCore.Load(path);
            foreach (Picture picture in dc.GetChildElements(true, ElementType.Picture))
            {
                if (imgInventory.Exists(img => img.GetStream().Length == picture.ImageData.GetStream().Length) == false)
                    imgInventory.Add(picture.ImageData);
            }

            if (imgInventory.Count == 0)
            {
                imgInventory.Clear();
                GC.Collect();
                return;
            }

            foreach (ImageData img in imgInventory)
            {
                using Image tempImage = Bitmap.FromStream(img.GetStream());
                imageTypes.Add(img.Format.ToString());

                result.Add(new Bitmap(tempImage));
            }

            imgInventory.Clear();
            GC.Collect();
        }
    }
}
