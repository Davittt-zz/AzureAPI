using HAA.Api.ActionFilters;
using System.Collections.Generic;
using System.Web.Http;
using HAA.BusinessServices.Base;
using HAA.BusinessEntities;
using System.Web.Routing;
using HAA.BusinessEntities.Input;

namespace HAA.Api.Controllers
{
    /// <summary>
    /// Controller for getting and saving speaker configuration per project
    /// </summary>
    //[AuthorizationRequired]
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
    }
}