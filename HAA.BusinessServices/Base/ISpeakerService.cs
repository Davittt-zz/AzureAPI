using HAA.BusinessEntities;
using HAA.BusinessEntities.Input;
using System.Collections.Generic;

namespace HAA.BusinessServices.Base
{
    public interface ISpeakerService
    {
        List<SpeakerConfigEntity> GetByProjectId(string projNumber);

        object CrossProduct(CrossProductModel model);

        float DotProduct(CrossProductModel model);

        float StandardDeviation(object array);

        float Radians(float input);

        float Degrees(float input);

        float TriangleArea(TriangleModel input);

        float PythagDist(PythagDistModel input);

        float Intersection(IntersectionModel input);

        object SpeakerAngles(SpeakerAnglesModel input);

        object DetermineWallType(DetermineWallTypeModel model);

        object CalcMirrorPnts(CalcMirrorPointsModel model);

        object CalcImagePoint(CalcImagePointModel model);

        object GetDolbyArray();

        object OrientSpkr(GetDolbyArrayModel input);

        object UpdateSpkrConfig(UpdateSpkrConfigModel input);
    }
}