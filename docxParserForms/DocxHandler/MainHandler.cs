using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json.Linq;
using static System.String;
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
                var json = sr.ReadToEnd();
                dynamic data = JObject.Parse(json);
                _connectionString = data.connectionString;
                _splitExample = data.splitExample;
            }
        }

        public List<Model> HandleFile(string filepath)
        {
            var (images, descriptions, models, imageTypes) 
                = (new List<Bitmap>(), new List<string>(), new List<Model>(), new List<string>());

            ImageHandler.ExtractImages(filepath, images, imageTypes);

            try
            {
                using (var wordDocument = WordprocessingDocument.Open(filepath, false))
                {
                    HandleParagraphsInBody(wordDocument, descriptions, images.Count);
                }

                CheckDescriptions(descriptions, images.Count);
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

        private void CheckDescriptions(IList<string> descriptions, int count)
        {
            if (descriptions.Count < count)
                for (var i = descriptions.Count; i < count; i++)
                    descriptions.Add("");
            else if (descriptions.Count > count)
            {
                for (var i = 0; i < count - descriptions.Count; i++)
                    descriptions.RemoveAt(-1);
            }
        }

        private void HandleParagraphsInBody(WordprocessingDocument wordDocument, List<string> descriptions, int imagesCount)
        {
            var body = wordDocument.MainDocumentPart.Document.Body;

            List<string> descriptionsInParagraph = new();
            string? description = null;
            var (paragraphCounter, tempImagesCount) = (0, 0);

            var isDescriptionContains = false;
            foreach (var paragraph in body.Descendants<Paragraph>())
            {
                paragraphCounter++;

                // переписать
                foreach (var run in paragraph.Descendants<Run>())
                {
                    if (descriptions.Count >= imagesCount)
                    {
                        ClearTempData(ref description, descriptionsInParagraph, ref paragraphCounter);
                        return;
                    }

                    var imagesCountInRun = run.Descendants<Drawing>().Count();
                    if (imagesCountInRun > 0)
                    {
                        if (!isDescriptionContains)
                        {
                            if (tempImagesCount == 0)
                            {
                                tempImagesCount = imagesCountInRun;
                                continue;
                            }

                            while (tempImagesCount != 0)
                            {
                                descriptions.Add("");
                                tempImagesCount--;
                            }

                            tempImagesCount = imagesCountInRun;
                            ClearTempData(ref description, descriptionsInParagraph, ref paragraphCounter);
                            continue;
                        }
                        else
                        {
                            if (Checker(descriptionsInParagraph, ref paragraphCounter, tempImagesCount, ref description, descriptions))
                                continue;
                        }
                    }
                    else
                    {
                        description = DescriptionHandler.GetDescription(paragraph);
                        if (description == null
                            || description?.Trim().Length == 0
                            || descriptionsInParagraph.Count != 0
                            && CompareOrdinal(description, descriptionsInParagraph[^1]) == 0)
                            continue;

                        descriptionsInParagraph.Add(description);
                        if (descriptionsInParagraph.Count == tempImagesCount)
                            break;
                        isDescriptionContains = true;
                        break;
                    }
                }

                if (Checker(descriptionsInParagraph, ref paragraphCounter, tempImagesCount, ref description, descriptions))
                    continue;
            }
        }

        private bool Checker(List<string> descriptionsInParagraph,
            ref int paragraphCounter, int tempImagesCount, ref string? description, List<string> descriptions)
        {
            if (CheckDscriptionsAndImages(descriptionsInParagraph, ref paragraphCounter, tempImagesCount))
                return true;

            if (AddNewLinesToLists(descriptions, descriptionsInParagraph) || paragraphCounter > 1)
                ClearTempData(ref description, descriptionsInParagraph, ref paragraphCounter);
            return false;
        }

        private bool CheckDscriptionsAndImages(List<string> descriptionsInParagraph,
            ref int paragraphCounter, int tempImagesCount)
        {
            descriptionsInParagraph = descriptionsInParagraph.Where(
                    description => description != null).ToList();

            if (descriptionsInParagraph.Count == 0)
                return true;

            if (tempImagesCount == 0)
            {
                paragraphCounter = 0;
                descriptionsInParagraph.Clear();
                return true;
            }

            if (descriptionsInParagraph.Count < tempImagesCount)
                DescriptionHandler.HandleDescriptionToImages(descriptionsInParagraph, _splitExample);

            return false;
        }

        private void WriteDataInModelsList(List<Bitmap> images, List<string> descriptions,
            List<Model> models, string path, List<string> imageTypes)
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

        private void ClearTempData(ref string? description, List<string> descriptionsInParagraph,
                ref int paragraphCounter)
        {
            description = null;
            descriptionsInParagraph.Clear();
            paragraphCounter = 0;
        }

        private bool AddNewLinesToLists(List<string> descriptions, List<string> tempDescriptions)
        {
            foreach (var description in tempDescriptions)
                descriptions.Add(description);

            return true;
        }
    }
}