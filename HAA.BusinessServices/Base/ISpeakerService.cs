using HAA.BusinessEntities;
using System.Collections.Generic;

namespace HAA.BusinessServices.Base
{
    public interface ISpeakerService
    {
       List<SpeakerConfigEntity> GetByProjectId(int projectId);
    }
}