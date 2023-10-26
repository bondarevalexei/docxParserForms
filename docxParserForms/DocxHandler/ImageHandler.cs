using SautinSoft.Document;
using SautinSoft.Document.Drawing;

namespace docxParserForms.DocxHandler
{
    public static class ImageHandler
    {
        public static void ExtractImages(string path, List<Bitmap> result, List<string> imageTypes)
        {
            var imgInventory = new List<ImageData>();

            var dc = DocumentCore.Load(path);
            imgInventory.AddRange(from Picture picture in dc.GetChildElements(true, ElementType.Picture)
                                  where imgInventory.Exists(img => img.GetStream().Length == picture.ImageData.GetStream().Length) == false
                                  select picture.ImageData);
            
            if (imgInventory.Count == 0)
                return;

            foreach (var img in imgInventory)
            {
                using var tempImage = Image.FromStream(img.GetStream());
                imageTypes.Add(img.Format.ToString());

                result.Add(new Bitmap(tempImage));
            }

            imgInventory.Clear();
        }
    }
}
