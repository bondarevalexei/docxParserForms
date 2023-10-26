using System.Collections;
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

        private void HandleParagraphsInBody(WordprocessingDocument wordDocument, ICollection<string> descriptions, int imagesCount)
        {
            var body = wordDocument.MainDocumentPart.Document.Body;

            List<string> descriptionsInParagraph = new();
            string? description = null;
            var (paragraphCounter, tempImagesCount, handledImagesCount) = (0, 0, 0);

            var isDescriptionContains = false;
            foreach (var paragraph in body.Descendants<Paragraph>())
            {
                paragraphCounter++;

                foreach (var run in paragraph.Descendants<Run>())
                {
                    handledImagesCount = descriptions.Count;

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

                            while (descriptions.Count != handledImagesCount + tempImagesCount)
                            {
                                descriptions.Add("");
                            }

                            tempImagesCount = imagesCountInRun;
                            ClearTempData(ref description, descriptionsInParagraph, ref paragraphCounter);
                        }
                        else
                        {
                            Checker(descriptionsInParagraph, ref paragraphCounter, tempImagesCount, ref description,
                                descriptions);
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

                Checker(descriptionsInParagraph, ref paragraphCounter, tempImagesCount, ref description, descriptions);
            }
        }

        private void Checker(List<string> descriptionsInParagraph,
            ref int paragraphCounter, int tempImagesCount, ref string? description, ICollection<string> descriptions)
        {
            if (CheckDescriptionsAndImages(descriptionsInParagraph, ref paragraphCounter, tempImagesCount)) return;

            if (AddNewLinesToLists(descriptions, descriptionsInParagraph) || paragraphCounter > 1)
                ClearTempData(ref description, descriptionsInParagraph, ref paragraphCounter);
        }

        private bool CheckDescriptionsAndImages(List<string> descriptionsInParagraph,
            ref int paragraphCounter, int tempImagesCount)
        {
            descriptionsInParagraph = descriptionsInParagraph.Where(_ => true).ToList();

            if (descriptionsInParagraph.Count == 0) return true;

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

        private void WriteDataInModelsList(IReadOnlyList<Bitmap> images, IReadOnlyList<string> descriptions,
            ICollection<Model> models, string path, IReadOnlyList<string> imageTypes)
        {
            for (var i = 0; i < descriptions.Count; i++)
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

        private void ClearTempData(ref string? description, IList descriptionsInParagraph,
                ref int paragraphCounter)
        {
            description = null;
            descriptionsInParagraph.Clear();
            paragraphCounter = 0;
        }

        private bool AddNewLinesToLists(ICollection<string> descriptions, List<string> tempDescriptions)
        {
            foreach (var description in tempDescriptions)
                descriptions.Add(description);

            return true;
        }
    }
}