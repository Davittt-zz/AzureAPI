using AutoMapper;
using HAA.BusinessEntities;
using HAA.BusinessEntities.Input;
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
        private HAA.SpeakerConfig.Functions _vbFunctions;
        /// <summary>
        /// Public constructor with DI Unity
        /// </summary>
        public SpeakerService(UnitOfWork unitOfWork)
        {
            _unitOfWork = unitOfWork;
            _vbFunctions = new SpeakerConfig.Functions();
        }

        public List<SpeakerConfigEntity> GetByProjectId(string projNumber)
        {
            var speakerConfigs = _unitOfWork.SpeakerConfigRepository.GetWithInclude(x => x.ProjectNumber == projNumber, "Project").ToList();
            if (speakerConfigs != null)
            {
                var config = new MapperConfiguration(cfg =>
                {
                    cfg.CreateMap<HAA.DataModel.SpeakerConfig, SpeakerConfigEntity>();
                    cfg.CreateMap<Project, ProjectEntity>();
                });

                IMapper mapper = config.CreateMapper();
                var speakerModel = mapper.Map<List<SpeakerConfigEntity>>(speakerConfigs);
                return speakerModel;
            }

            return null;
        }

        public object CrossProduct(CrossProductModel model)
        {
            return _vbFunctions.CrossProduct(model.V1.X, model.V1.Y, model.V1.Z, model.V2.X, model.V2.Y,
                model.V2.Z, model.V3.X, model.V3.Y, model.V3.Z);
        }
    }
}
