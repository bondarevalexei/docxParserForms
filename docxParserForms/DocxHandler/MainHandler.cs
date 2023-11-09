using Newtonsoft.Json.Linq;
using SautinSoft.Document;
using Paragraph = SautinSoft.Document.Paragraph;
using System.Linq;

namespace docxParserForms.DocxHandler
{
    public class MainHandler
    {
        private readonly string _connectionString, _splitExample;
        private readonly int _descriptionsIndexForHash, _minImageWidth, _minImageHeight;
        private readonly bool _isPatent;

        public MainHandler()
        {
            using StreamReader sr = new("./../../../appsettings.json");
            var json = sr.ReadToEnd();
            dynamic data = JObject.Parse(json);
            (_connectionString, _splitExample, _descriptionsIndexForHash, _minImageWidth, _minImageHeight, _isPatent)
                = (data.connectionString, data.splitExample, data.descriptionsIndexForHash, data.minImageWidth, data.minImageHeight, data.isPatent);
        }

        public List<Model> HandleFile(string filepath)
        {
            var (images, descriptions, models, imageTypes)
                = (new List<Bitmap>(), new List<string>(), new List<Model>(), new List<string>());

            try
            {
                ImageHandler.ExtractImages(filepath, images, imageTypes);
                var descriptionsHash = TextHandler.CollectDescriptionsFromText(filepath);
                HandleParagraphsInBody(descriptions, descriptionsHash, filepath);
                CheckDescriptions(descriptions, images.Count);
                WriteDataInModelsList(images, descriptions, models, filepath, imageTypes);
                descriptionsHash.Clear();
                images.Clear(); descriptions.Clear(); imageTypes.Clear();

                //DbHandler.SaveToDb(descriptions, images, _connectionString);
                MessageBox.Show($"Файл {filepath} успешно обработан. Добавлено {models.Count} элемента(ов).");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            return models;
        }

        private void CheckDescriptions(IList<string> descriptions, int imageCount)
        {
            while (descriptions.Count < imageCount)
                descriptions.Add("");
        }

        private void HandleParagraphsInBody(ICollection<string> descriptions,
            SortedDictionary<string, string> descriptionsHash, string filePath)
        {
            var dc = DocumentCore.Load(filePath);
            List<int> imagesCountInRuns = new();

            foreach (var element in dc.GetChildElements(true, ElementType.Paragraph))
            {
                var paragraph = (Paragraph)element;
                    var imagesCountInParagraph = CountImagesForParagraph(paragraph);
                    if (imagesCountInParagraph > 0) 
                        imagesCountInRuns.Add(imagesCountInParagraph);
            }

            CompareImagesWithDescriptions(descriptions, imagesCountInRuns, descriptionsHash);
        }

        private void CompareImagesWithDescriptions(ICollection<string> descriptions, List<int> imagesCountInRuns, SortedDictionary<string, string> descriptionsHash)
        {
            for(int i = 0; i < imagesCountInRuns.Count; i++)
            {
                var key = FindDescription(descriptionsHash, i);
                if (key != null) {
                    if (imagesCountInRuns[i] > 1)
                    {
                        if (DescriptionHandler.IsDescriptionDevided(descriptionsHash.GetValueOrDefault(key), _splitExample))
                        {
                            var tempDescr = DescriptionHandler.HandleDescriptionToMany(descriptionsHash.GetValueOrDefault(key), _splitExample);
                            for (int j = 0; j < Math.Min(tempDescr.Count, imagesCountInRuns[i]); j++)
                            {
                                descriptions.Add(tempDescr[j]);
                            }

                            while (tempDescr.Count < imagesCountInRuns[i])
                            {
                                descriptions.Add("");
                            }
                        }
                    }
                    else
                        descriptions.Add(descriptionsHash.GetValueOrDefault(key));
                }
                else
                {
                    descriptions.Add("");
                }
            }
        }

        private string? FindDescription(SortedDictionary<string, string> descriptionsHash, int i)
        {
            foreach(var key in descriptionsHash.Keys)
            {
                double temp;
                double.TryParse(key.Split(' ')[1], out temp);

                if ((int)temp == i + 1) return key;
            }

            return null;
        }

        private int CountImagesForParagraph(Element el)
        {
            var count = 0;
            Paragraph par = (Paragraph)el;
            count += (from picture in par.GetChildElements(true, ElementType.Picture)
                      select picture).Count();
            return count;
        }

        private void WriteDataInModelsList(IList<Bitmap> images, IList<string> descriptions,
            ICollection<Model> models, string path, IList<string> imageTypes)
        {
            for (var i = 0; i < descriptions.Count; i++)
            {
                if (images[i].Width >= _minImageWidth && images[i].Height >= _minImageHeight)
                {
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
                else
                {
                    images.RemoveAt(i);
                    descriptions.RemoveAt(i);
                    imageTypes.RemoveAt(i);
                }
            }
        }
    }
}