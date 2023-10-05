using docxParserForms.Interfaces;
using docxParserForms.Models;
using Microsoft.EntityFrameworkCore;

namespace docxParserForms.Db.Context
{
    public sealed class ModelsDbContext : DbContext, IModelsDbContext
    {
        public DbSet<Model> Models { get; set; }

        public ModelsDbContext(DbContextOptions<ModelsDbContext> options) : base(options) { }
    }
}
