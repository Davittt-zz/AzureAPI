using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;

namespace HAA.DataModel.Repositories
{
    public class SpeakerRepository
    {
        internal WebApiDbEntities Context;
        internal DbSet<SpeakerConfig> DbSet;

        public SpeakerRepository(WebApiDbEntities context)
        {
            this.Context = context;
            this.DbSet = context.Set<SpeakerConfig>();
        }

        public virtual List<SpeakerConfig> GetByProjectId(int id)
        {
            return DbSet.Where(x => x.ProjectId == id).ToList();
        }
    }
}