using docxParserForms.Db.Context;

namespace docxParserForms.Db
{
    public class DbInitializer
    {
        public static void Initialize(ModelsDbContext context)
        {
            context.Database.EnsureCreated();
        }
    }
}
