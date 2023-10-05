using docxParserForms.Models;
using Microsoft.EntityFrameworkCore;

namespace docxParserForms.Interfaces
{
    public interface IModelsDbContext
    {
        DbSet<Model> Models { get; set; }
        Task<int> SaveChangesAsync(CancellationToken cancellationToken);
    }
}
