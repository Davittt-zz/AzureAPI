using HAA.BusinessEntities;
using HAA.BusinessEntities.Input;
using System.Collections.Generic;

namespace HAA.BusinessServices.Base
{
    public interface ISpeakerService
    {
        List<SpeakerConfigEntity> GetByProjectId(string projNumber);

        object CrossProduct(CrossProductModel model);
    }
}