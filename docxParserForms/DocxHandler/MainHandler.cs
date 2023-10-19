using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json.Linq;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;

namespace docxParserForms.DocxHandler
{
    public class MainHandler
    {
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
            List<string> imageTypes = new();

            try
            {
                using (WordprocessingDocument wordDocument =
                    WordprocessingDocument.Open(filepath, false))
                {
                    HandleParagraphsInBody(wordDocument, images, descriptions, imageTypes);
                }

                WriteDataInModelsList(images, descriptions, models, filepath, imageTypes);

                //DbHandler.SaveToDb(descriptions, images, _connectionString);
                MessageBox.Show($"Файл {filepath} успешно обработан. Добавлено {descriptions.Count} элемента(ов).");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            return models;
        }

        private void WriteDataInModelsList(List<Bitmap> images,
            List<string> descriptions, List<Model> models, string path,
                List<string> imageTypes)
        {
            for (int i = 0; i < descriptions.Count; i++)
                models.Add(new Model
                {
                    Description = descriptions[i],
                    Image = images[i],
                    Filename = path.Split('\\')[^1],
                    ImageFormat = imageTypes[i],
                    Width = images[i].Width,
                    Height = images[i].Height,
                });
        }

        private void HandleParagraphsInBody(WordprocessingDocument wordDocument,
            List<Bitmap> images, List<string> descriptions, List<string> imageTypes)
        {
            Body body = wordDocument.MainDocumentPart.Document.Body;

            List<Bitmap> imagesInPargraph = new();
            List<string> descriptionsInParagraph = new();

            Bitmap? imageBitmap = null;
            string? description = null;
            int paragraphCounter = 0, imageFlag = 0;

            foreach (Paragraph paragraph in body.Descendants<Paragraph>())
            {
                paragraphCounter++;

                foreach (Run run in paragraph.Descendants<Run>())
                {
                    Drawing image =
                        run.Descendants<Drawing>().FirstOrDefault();

                    if (image != null && image.Inline != null)

                    {
                        if (imagesInPargraph.Count > 0 && paragraphCounter > imageFlag)
                        {
                            descriptionsInParagraph.Add("");
                            AddNewLinesToLists(images, descriptions, imagesInPargraph, 
                                descriptionsInParagraph);
                            ClearTempData(ref imageBitmap, ref description, imagesInPargraph,
                                descriptionsInParagraph, ref paragraphCounter, ref imageFlag);
                        }

                        imageBitmap = new Bitmap(ImageHandler.ExtractImage(wordDocument.MainDocumentPart, 
                            image, imageTypes));
                        imagesInPargraph.Add(imageBitmap);
                        imageFlag = paragraphCounter;
                    }
                    else
                    {
                        description = DescriptionHandler.GetDescription(paragraph);
                        if (description == null
                            || description?.Trim().Length == 0
                            || descriptionsInParagraph.Count != 0
                            && String.Compare(description, descriptionsInParagraph[^1]) == 0)
                            continue;

                        descriptionsInParagraph.Add(description);
                        break;
                    }
                }

                if (CheckDscriptionsAndImages(descriptionsInParagraph, imagesInPargraph, ref paragraphCounter))
                    continue;

                if (AddNewLinesToLists(images, descriptions, imagesInPargraph, descriptionsInParagraph)
                    || paragraphCounter > 1)
                    ClearTempData(ref imageBitmap, ref description, imagesInPargraph,
                                    descriptionsInParagraph, ref paragraphCounter, ref imageFlag);
            }
        }

        private bool CheckDscriptionsAndImages(List<string> descriptionsInParagraph, 
            List<Bitmap> imagesInPargraph, ref int paragraphCounter)
        {
            descriptionsInParagraph = descriptionsInParagraph.Where(
                    description => description != null).ToList();

            if (descriptionsInParagraph.Count == 0)
                return true;

            if (imagesInPargraph.Count == 0)
            {
                paragraphCounter = 0;
                descriptionsInParagraph.Clear();
                return true;
            }

            if (descriptionsInParagraph.Count < imagesInPargraph.Count)
                DescriptionHandler.HandleDescriptionToImages(descriptionsInParagraph, _splitExample);

            return false;
        }

        private void ClearTempData(ref Bitmap? imageBitmap, ref string? description,
            List<Bitmap> imagesInParagraph, List<string> descriptionsInParagraph,
                ref int paragraphCounter, ref int imageFlag)
        {
            (imageBitmap, description) = (null, null);
            imagesInParagraph.Clear();
            descriptionsInParagraph.Clear();
            (paragraphCounter, imageFlag) = (0, 0);
        }

        private bool AddNewLinesToLists(List<Bitmap> images,
            List<string> descriptions, List<Bitmap> tempImages, List<string> tempDescriptions)
        {
            for (int i = 0; i < descriptions.Count; i++)
            {
                descriptions.Add(tempDescriptions[i]);
                images.Add(tempImages[i]);
            }
            return true;
        }
    }
}