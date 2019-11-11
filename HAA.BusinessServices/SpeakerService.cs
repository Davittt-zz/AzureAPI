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

        public List<SpeakerConfigEntity> GetByProjectId(string projNumber)
        {
            var speakerConfigs = _unitOfWork.SpeakerConfigRepository.GetWithInclude(x => x.ProjectNumber == projNumber, "Project").ToList();
            if (speakerConfigs != null)
            {
                var config = new MapperConfiguration(cfg =>
                {
                    cfg.CreateMap<SpeakerConfig, SpeakerConfigEntity>();
                    cfg.CreateMap<Project, ProjectEntity>();
                });

                IMapper mapper = config.CreateMapper();
                var speakerModel = mapper.Map<List<SpeakerConfigEntity>>(speakerConfigs);
                return speakerModel;
            }

            return null;
        }
    }
}
