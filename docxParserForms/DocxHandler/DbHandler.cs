using System.Data.SqlClient;
using System.Drawing.Imaging;

namespace docxParserForms.DocxHandler
{
    public static class DbHandler
    {
        public static void SaveToDb(List<Model> models, string _connectionString)
        {
            foreach (var model in models)
            {
                using var connection = new SqlConnection(_connectionString);
                using var command = connection.CreateCommand();
                command.CommandText =
                    "INSERT INTO parser_db (description, image) VALUES (@description, @image)";

                using var memoryStream = new MemoryStream();
                {
                    model.Image.Save(memoryStream, ImageFormat.Jpeg);
                    memoryStream.Position = 0;

                    var sqlImageParam = new SqlParameter(
                        "@image", System.Data.SqlDbType.VarBinary, (int)memoryStream.Length)
                    {
                        Value = memoryStream.ToArray()
                    };

                    var sqlDescriptionParam = new SqlParameter(
                        "@description", System.Data.SqlDbType.VarChar, model.Description.Length)
                    {
                        Value = model.Description
                    };

                    var sqlFilenameParam = new SqlParameter(
                        "@filename", System.Data.SqlDbType.VarChar, model.Filename.Length)
                    {
                        Value = model.Filename
                    };

                    var sqlWidthParam = new SqlParameter("@width", System.Data.SqlDbType.Int, sizeof(int))
                    {
                        Value = model.Width
                    };

                    var sqlHeightParam = new SqlParameter("@width", System.Data.SqlDbType.Int, sizeof(int))
                    {
                        Value = model.Height
                    };

                    var sqlImgFormatParam = new SqlParameter(
                        "@img_format", System.Data.SqlDbType.VarChar, model.ImageFormat.Length)
                    {
                        Value = model.ImageFormat
                    };

                    var sqlGraphNameParam = new SqlParameter(
                        "@filename", System.Data.SqlDbType.VarChar, model.GrahicsType.Length)
                    {
                        Value = model.GrahicsType
                    };

                    var sqlIsHumanParam = new SqlParameter(
                        "@is_human", System.Data.SqlDbType.Bit, sizeof(bool))
                    {
                        Value = model.IsAstronautOrPilot ? 1 : 0
                    };

                    var sqlIsFormulaParam = new SqlParameter(
                        "@is_formula", System.Data.SqlDbType.Bit, sizeof(bool))
                    {
                        Value = model.IsFormula ? 1 : 0
                    };

                    command.Parameters.Add(sqlDescriptionParam);
                    command.Parameters.Add(sqlImageParam);
                    command.Parameters.Add(sqlFilenameParam);
                    command.Parameters.Add(sqlWidthParam);
                    command.Parameters.Add(sqlHeightParam);
                    command.Parameters.Add(sqlImgFormatParam);
                    command.Parameters.Add(sqlGraphNameParam);
                    command.Parameters.Add(sqlIsHumanParam);
                    command.Parameters.Add(sqlIsFormulaParam);
                }
                connection.Open();
                command.ExecuteNonQuery();
            }
        }
    }
}
