using BusinessServices.Base;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BusinessEntities;
using DataModel.UnitOfWork;
using DataModel;
using AutoMapper;
using System.Transactions;

namespace BusinessServices
{
	public class ElementServices : IElementServices
	{
		private readonly UnitOfWork _unitOfWork;

		/// <summary>
		/// Public constructor with DI Unity
		/// </summary>
		public ElementServices(UnitOfWork unitOfWork)
		{
			_unitOfWork = unitOfWork;
		}


		/// <summary>
		/// Fetches element details by id
		/// </summary>
		/// <param name="elementId"></param>
		/// <returns></returns>
		public BusinessEntities.ElementEntity GetElementById(int elementId)
		{
			var element = _unitOfWork.ElementRepository.GetByID(elementId);
			if (element != null)
			{
				var config = new MapperConfiguration(cfg =>
				{
					cfg.CreateMap<Element, ElementEntity>();
				});

				IMapper mapper = config.CreateMapper();
				var elementsModel = mapper.Map<Element, ElementEntity>(element);
				return elementsModel;
			}
			return null;
		}

		/// <summary>
		/// Fetches all the elements.
		/// </summary>
		/// <returns></returns>
		public IEnumerable<BusinessEntities.ElementEntity> GetAllElements()
		{
			var elements = _unitOfWork.ElementRepository.GetAll().ToList();
			if (elements.Any())
			{
				var config = new MapperConfiguration(cfg =>
				{
					cfg.CreateMap<Element, ElementEntity>();
				});
				IMapper mapper = config.CreateMapper();
				
				var elementsModel = mapper.Map<List<Element>, List<ElementEntity>>(elements);
				return elementsModel;
			}
			return null;
		}

		/// <summary>
		/// Creates a element
		/// </summary>
		/// <param name="ElementEntity"></param>
		/// <returns></returns>
		public int CreateElement(BusinessEntities.ElementEntity ElementEntity)
		{
			using (var scope = new TransactionScope(TransactionScopeOption.Required))
			{
				var element = new Element
				{
					Name = ElementEntity.Name,
					Created = ElementEntity.Created,
					Modified = ElementEntity.Modified
				};
				_unitOfWork.ElementRepository.Insert(element);
				_unitOfWork.Save();
				scope.Complete();
				return element.Id;
			}
		}

		/// <summary>
		/// Updates a element
		/// </summary>
		/// <param name="elementId"></param>
		/// <param name="ElementEntity"></param>
		/// <returns></returns>
		public bool UpdateElement(int elementId, BusinessEntities.ElementEntity ElementEntity)
		{
			var success = false;
			if (ElementEntity != null)
			{
				using (var scope = new TransactionScope())
				{
					var element = _unitOfWork.ElementRepository.GetByID(elementId);
					if (element != null)
					{
						element.Name = ElementEntity.Name;
						_unitOfWork.ElementRepository.Update(element);
						_unitOfWork.Save();
						scope.Complete();
						success = true;
					}
				}
			}
			return success;
		}

		/// <summary>
		/// Deletes a particular element
		/// </summary>
		/// <param name="elementId"></param>
		/// <returns></returns>
		public bool DeleteElement(int elementId)
		{
			var success = false;
			if (elementId > 0)
			{
				using (var scope = new TransactionScope())
				{
					var element = _unitOfWork.ElementRepository.GetByID(elementId);
					if (element != null)
					{
						_unitOfWork.ElementRepository.Delete(element);
						_unitOfWork.Save();
						scope.Complete();
						success = true;
					}
				}
			}
			return success;
		}
	}
}