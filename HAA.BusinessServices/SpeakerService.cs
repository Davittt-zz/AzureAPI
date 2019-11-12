using AutoMapper;
using HAA.BusinessEntities;
using HAA.BusinessEntities.Input;
using HAA.BusinessServices.Base;
using HAA.DataModel;
using HAA.DataModel.UnitOfWork;
using System;
using System.Collections.Generic;
using System.Linq;

namespace HAA.BusinessServices
{
    public class SpeakerService : ISpeakerService
    {
        private readonly UnitOfWork _unitOfWork;
        private readonly SpeakerConfig.Functions _vbFunctions;
        private readonly SpeakerConfig.Library _vbLibrary;
        private readonly SpeakerConfig.DataFunctions _vbDataFunctions;
        private readonly SpeakerConfig.SpeakerConfig _vbSpeakerConfig;

        /// <summary>
        /// Public constructor with DI Unity
        /// </summary>
        public SpeakerService(UnitOfWork unitOfWork)
        {
            _unitOfWork = unitOfWork;
            _vbFunctions = new SpeakerConfig.Functions();
            _vbLibrary = new SpeakerConfig.Library();
            _vbDataFunctions = new SpeakerConfig.DataFunctions();
            _vbSpeakerConfig = new SpeakerConfig.SpeakerConfig();
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

        public float DotProduct(CrossProductModel model)
        {
            return _vbFunctions.DotProduct(model.V1.X, model.V1.Y, model.V1.Z, model.V2.X, model.V2.Y,
               model.V2.Z, model.V3.X, model.V3.Y, model.V3.Z);
        }

        public float StandardDeviation(object array)
        {
            return _vbFunctions.StandardDeviation(ref array);
        }

        public float Radians(float input)
        {
            var resultObj = _vbFunctions.Radians(input);
            var tryParse = float.TryParse(resultObj.ToString(), out float result);
            if (!tryParse) throw new ArgumentException();
            return result;
        }

        public float Degrees(float input)
        {
            var resultObj = _vbFunctions.Degrees(input);
            var tryParse = float.TryParse(resultObj.ToString(), out float result);
            if (!tryParse) throw new ArgumentException();
            return result;
        }

        public float TriangleArea(TriangleModel input)
        {
            object p1 = input.P1;
            object p2 = input.P2;
            object p3 = input.P3;
            var resultObj = _vbFunctions.TriangleArea(ref p1, ref p2, ref p3);
            var tryParse = float.TryParse(resultObj.ToString(), out float result);
            if (!tryParse) throw new ArgumentException();
            return result;
        }

        public float PythagDist(PythagDistModel input)
        {
            return _vbFunctions.PythagDist(input.X1, input.X2, input.Y1, input.Y2);
        }

        public float Intersection(IntersectionModel input)
        {
            var p1 = input.P1;
            var p2 = input.P2;
            var corner1 = input.Corner1;
            var corner2 = input.Corner2;
            var corner3 = input.Corner3;
            var corner4 = input.Corner4;
            var resultObj = _vbFunctions.Intersection(ref p1, ref p2, ref corner1, ref corner2, ref corner3, ref corner4);
            var tryParse = float.TryParse(resultObj.ToString(), out float result);
            if (!tryParse) throw new ArgumentException();
            return result;
        }

        public object SpeakerAngles(SpeakerAnglesModel input)
        {
            var config = input.Config;
            var brand = input.Brand;
            return _vbLibrary.SpeakerAngles(ref config, ref brand);
        }

        public object DetermineWallType(DetermineWallTypeModel model)
        {
            var coordin = model.Coordin;
            var label = model.SpeakerLabel;
            return _vbDataFunctions.DetermineWallType(ref coordin, ref label);
        }

        public object CalcImagePoint(CalcImagePointModel model)
        {
            var mainLp = model.MainLP;
            var speakerArr = model.SpeakerArr;
            var spkrPick = model.SpkrPick;
            var topPnt = model.TopPnt;
            return _vbDataFunctions.CalcImagePoint(ref mainLp, ref speakerArr, ref spkrPick, ref topPnt);
        }

        public object CalcMirrorPnts(CalcMirrorPointsModel model)
        {
            var mainLp = model.MainLP;
            var speakerCnt = model.SpeakerCnt;
            var speakerInfo = model.SpeakerInfo;
            var speakerConfig = model.SpkrConfig;
            var speakerImage = model.SpkrImage;
            var topWidth = model.TopWidth;
            var wallArr = model.WallArr;

            return _vbDataFunctions.CalcMirrorPnts(ref mainLp, ref wallArr, ref speakerImage, ref speakerConfig, ref topWidth, ref speakerCnt, ref speakerInfo);
        }

        public object GetDolbyArray()
        {
            return _vbSpeakerConfig.GetDolbyArray();
        }

        public object OrientSpkr(GetDolbyArrayModel input)
        {
            var projNum = input.ProjNum;
            var spkrLstIn = input.SpkrLstIn;
            return _vbSpeakerConfig.OrientSpkr(ref projNum, ref spkrLstIn);
        }

        public object UpdateSpkrConfig(UpdateSpkrConfigModel input)
        {
            var projNum = input.ProjNumb;
            var configTypeDolby = input.ConfigTypeDolby;
            var seatss = input.Seatss;
            var selectedSpkrConfigur = input.SelectedSpkrConfigur;
            var spkrInfo = input.SpkrInfo;

            return _vbSpeakerConfig.UpdateSpkrConfig(ref projNum, ref selectedSpkrConfigur, ref configTypeDolby, ref seatss,
            ref spkrInfo);
        }
    }
}