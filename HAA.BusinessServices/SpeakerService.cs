using AutoMapper;
using HAA.BusinessEntities;
using HAA.BusinessServices.Base;
using HAA.DataModel;
using HAA.DataModel.UnitOfWork;
using System.Collections.Generic;
using System.Linq;

namespace HAA.BusinessServices
{
    public class SpeakerService : ISpeakerService
    {
        private readonly UnitOfWork _unitOfWork;

        /// <summary>
        /// Public constructor with DI Unity
        /// </summary>
        public SpeakerService(UnitOfWork unitOfWork)
        {
            _unitOfWork = unitOfWork;
        }

        public List<SpeakerConfigEntity> GetByProjectId(int projectId)
        {
            var speakerConfigs = _unitOfWork.SpeakerConfigRepository.GetManyQueryable(x => x.ProjectId == projectId).ToList();
            if (speakerConfigs != null)
            {
                var config = new MapperConfiguration(cfg =>
                {
                    cfg.CreateMap<SpeakerConfig, SpeakerConfigEntity>();
                });

                IMapper mapper = config.CreateMapper();
                var speakerModel = mapper.Map<List<SpeakerConfig>, List<SpeakerConfigEntity>>(speakerConfigs);
                return speakerModel;
            }

            return null;
        }
    }
}
