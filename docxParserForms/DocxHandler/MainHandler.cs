using Newtonsoft.Json.Linq;
using SautinSoft.Document;
using System.Collections;
using System.Globalization;
using System.Text;
using Paragraph = SautinSoft.Document.Paragraph;

namespace docxParserForms.DocxHandler
{
    public class MainHandler
    {
        private readonly string _connectionString, _splitExample;
        private readonly int _minImageWidth, _minImageHeight;

        public MainHandler()
        {
            using StreamReader sr = new("./../../../appsettings.json");
            var json = sr.ReadToEnd();
            dynamic data = JObject.Parse(json);
            (_connectionString, _splitExample, _minImageWidth, _minImageHeight)
                = (data.connectionString, data.splitExample, data.minImageWidth, data.minImageHeight);
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

            while (descriptions.Count > imageCount)
                descriptions.RemoveAt(descriptions.Count - 1);
        }

        private void HandleParagraphsInBody(ICollection<string> descriptions,
            Hashtable descriptionsHash, string filePath)
        {
            var dc = DocumentCore.Load(filePath);
            var isDescNeed = false;
            List<string> descriptionsBefore = new();
            var (tempImagesCount, parAfter, isBefore) = (0, 0, false);

            foreach (var element in dc.GetChildElements(true, ElementType.Paragraph))
            {
                if (isDescNeed) parAfter++;

                var paragraph = (Paragraph)element;
                var childElements = paragraph.GetChildElements(true).ToList();

                var imagesCountInParagraph = CountImagesForParagraph(paragraph);
                if (imagesCountInParagraph > 0)
                {
                    if (isDescNeed)
                    {
                        while (descriptions.Count < tempImagesCount)
                        {
                            if (descriptionsBefore.Count != 0)
                            {
                                descriptions.Add(descriptionsBefore[0]);
                                descriptionsBefore.RemoveAt(0);
                            }
                            else
                            {
                                descriptions.Add("");
                            }
                        }

                        isDescNeed = false;
                        parAfter = 0;
                    }

                    tempImagesCount += imagesCountInParagraph;
                    CheckRuns(childElements, descriptionsHash, descriptions, isBefore, descriptionsBefore);
                }
                else
                {
                    if (isDescNeed)
                    {
                        if (parAfter <= 20)
                            CheckRuns(childElements, descriptionsHash, descriptions, isBefore, descriptionsBefore);
                        else
                        {
                            if (descriptionsBefore.Count != 0)
                            {
                                descriptions.Add(descriptionsBefore[0]);
                                descriptionsBefore.RemoveAt(0);
                            }
                            else
                            {
                                descriptions.Add("");
                            }

                            isDescNeed = false;
                            parAfter = 0;
                        }
                    }
                    else
                    {
                        isBefore = true;
                        CheckRuns(childElements, descriptionsHash, descriptions, isBefore, descriptionsBefore);
                    }
                }

                isDescNeed = descriptions.Count < tempImagesCount;
            }
        }

        private void CheckRuns(List<Element> childElements, Hashtable? descriptionsHash,
            ICollection<string> descriptions, bool isBefore, List<string> descriptionsBefore)
        {
            bool isKeyUsed = false;
            StringBuilder sb = new();
            var tempIndex = 0;

            foreach (var child in childElements)
            {
                if (child.ElementType == ElementType.Run)
                {
                    var run = (Run)child;

                    if (DescriptionHandler.IsOnlyKey(run.Text) || tempIndex < 3 && isKeyUsed)
                    {
                        isKeyUsed = true;
                        sb.Append(run.Text.Trim());
                        sb.Append(' ');
                        tempIndex++;
                        continue;
                    }

                    string? description = !isKeyUsed
                        ? DescriptionHandler.GetDescription(run.Text, descriptionsHash)
                        : DescriptionHandler.GetDescription(sb.ToString().Trim(), descriptionsHash);

                    AddDescription(description, !isBefore ? descriptions : descriptionsBefore);


                    isKeyUsed = false;
                    tempIndex = 0;
                    sb.Clear();
                }
            }

            if (isKeyUsed && sb.Length != 0)
            {
                string? description = DescriptionHandler.GetDescription(sb.ToString().Trim(), descriptionsHash);
                AddDescription(description, !isBefore ? descriptions : descriptionsBefore);
            }
        }

        private void AddDescription(string? description, ICollection<string> descriptions)
        {
            if (description != null)
            {
                if (DescriptionHandler.IsDescriptionDevided(description, _splitExample))
                {
                    var tempDescriptions = DescriptionHandler.HandleDescriptionToMany(description, _splitExample);
                    foreach (var item in tempDescriptions)
                        descriptions.Add(item);
                }
                else
                    descriptions.Add(description);
            }
        }

        private string? FindDescription(Hashtable descriptionsHash, int i)
        {
            foreach (string key in descriptionsHash.Keys)
            {
                IFormatProvider formatter = new NumberFormatInfo { NumberDecimalSeparator = "." };
                var style = NumberStyles.Number | NumberStyles.AllowCurrencySymbol;

                if (DescriptionHandler.CheckForNumber(key.Split(' ')[1], style, formatter))
                    return key;

                _ = double.TryParse(key.Split(' ')[1], out double temp);
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