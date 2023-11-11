using DocxParserForms;

namespace docxParserForms
{
    public static class GraphicsClassificationResult
    {
        public static (bool, string) GetClassificatorResult(Bitmap image)
        {
            ImageConverter converter = new();

            var sampleData = new GraphicsClassificator.ModelInput
            {
                ImageSource = converter.ConvertTo(image, typeof(byte[])) as byte[]
            };

            var result = GraphicsClassificator.Predict(sampleData);
            var scoreResult = result.Score.ToList();
            scoreResult.Sort((a, b) =>
            {
                if (b > a) return 1;
                if (Math.Abs(b - a) < 1e-5) return 0;
                return -1;
            });

            return (scoreResult[0] + scoreResult[1] >= 0.85, result.PredictedLabel);
        }
    }
}
