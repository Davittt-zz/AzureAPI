using HAA.Api.ActionFilters;
using System.Collections.Generic;
using System.Web.Http;
using HAA.BusinessServices.Base;
using HAA.BusinessEntities;

namespace HAA.Api.Controllers
{
    /// <summary>
    /// Controller for getting and saving speaker configuration per project
    /// </summary>
    [AuthorizationRequired]
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
        /// <param name="projectId">The Project Id</param>
        /// <returns>A SpeakerConfig model</returns>
        [Route("speakerConfig/{projectId}")]
        public List<SpeakerConfigEntity> GetSpeakerConfigs(int projectId)
        {
            return _speakerService.GetByProjectId(projectId);
        }
    }
}