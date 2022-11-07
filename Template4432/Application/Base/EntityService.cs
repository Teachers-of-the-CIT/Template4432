using System.Data.Entity;
using System.Linq;
using Template4432.Contexts;
using Template4432.Models.Base;

namespace Template4432.Application.Base
{
    public abstract class EntityService<TEntity>
        where TEntity : Entity
    {
        protected ApplicationContext _context;
        protected DbSet<TEntity> _dbSet;

        protected EntityService(ApplicationContext context)
        {
            _context = context;
            _dbSet = context.Set<TEntity>();
        }

        public virtual bool Create(TEntity entity)
        {
            try
            {
                _dbSet.Add(entity);

                _context.SaveChanges();
            }
            catch
            {
                return false;
            }
            
            return true;
        }

        public virtual IQueryable<TEntity> ReadAsQueryable()
        {
            return _dbSet.AsQueryable();
        }
    }
}