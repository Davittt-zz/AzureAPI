using HAA.Api.ActionFilters;
using HAA.BusinessEntities;
using HAA.BusinessEntities.Input;
using HAA.BusinessServices.Base;
using System.Collections.Generic;
using System.Web.Http;
using System.Web.Routing;

namespace HAA.Api.Controllers
{
    /// <summary>
    /// Controller for getting and saving speaker configuration per project
    /// </summary>
    [AuthorizationRequired]
    [RoutePrefix("api/speakerConfig")]
    public class SpeakerConfigController : ApiController
    {
        private readonly ISpeakerService _speakerService;

        /// <summary>
        /// Constructor method
        /// </summary>
        /// <param name="speakerService"></param>
        public SpeakerConfigController(ISpeakerService speakerService)
        {
            _speakerService = speakerService;
        }

        /// <summary>
        /// Obtains the speaker settings given a project id.
        /// </summary>
        /// <param name="projNumber">The Project Id</param>
        /// <returns>A SpeakerConfig model</returns>
        [Route("{projNumber}")]
        public List<SpeakerConfigEntity> GetSpeakerConfigs(string projNumber)
        {
            return _speakerService.GetByProjectId(projNumber);
        }

        [Route("crossProduct")]
        [HttpPost]
        public object CrossProduct(CrossProductModel model)
        {
            return _speakerService.CrossProduct(model);
        }

        [Route("dotProduct")]
        [HttpPost]
        public object DotProduct(CrossProductModel model)
        {
            return _speakerService.DotProduct(model);
        }

        [Route("standardDeviation")]
        [HttpPost]
        public float StandardDeviation(object array)
        {
            return _speakerService.StandardDeviation(array);
        }

        [Route("radians/{input}")]
        public float Radians(float input)
        {
            return _speakerService.Radians(input);
        }

        [Route("degrees/{input}")]
        public float Degrees(float input)
        {
            return _speakerService.Degrees(input);
        }

        [Route("triangleArea")]
        [HttpPost]
        public float TriangleArea(TriangleModel input)
        {
            return _speakerService.TriangleArea(input);
        }

        [Route("pythagDist")]
        [HttpPost]
        public float PythagDist(PythagDistModel input)
        {
            return _speakerService.PythagDist(input);
        }

        [Route("intersection")]
        [HttpPost]
        public float Intersection(IntersectionModel input)
        {
            return _speakerService.Intersection(input);
        }

        [Route("speakerAngles")]
        [HttpPost]
        public object SpeakerAngles(SpeakerAnglesModel input)
        {
            return _speakerService.SpeakerAngles(input);
        }

        [Route("determineWallType")]
        [HttpPost]
        public object DetermineWallType(DetermineWallTypeModel input)
        {
            return _speakerService.DetermineWallType(input);
        }

        [Route("calcMirrorPnts")]
        [HttpPost]
        public object CalcMirrorPnts(CalcMirrorPointsModel input)
        {
            return _speakerService.CalcMirrorPnts(input);
        }

        [Route("calcImagePoint")]
        [HttpPost]
        public object CalcImagePoint(CalcImagePointModel input)
        {
            return _speakerService.CalcImagePoint(input);
        }

        [Route("dolbyArray")]
        public object GetDolbyArray()
        {
            return _speakerService.GetDolbyArray();
        }

        [Route("orientSpkr")]
        [HttpPost]
        public object OrientSpkr(GetDolbyArrayModel input)
        {
            return _speakerService.OrientSpkr(input);
        }

        [Route("updateSpkrConfig")]
        [HttpPost]
        public object UpdateSpkrConfig(UpdateSpkrConfigModel input)
        {
            return _speakerService.UpdateSpkrConfig(input);
        }
    }
}