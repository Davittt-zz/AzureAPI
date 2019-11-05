using HAA.BusinessEntities;
using System.Collections.Generic;


namespace HAA.BusinessServices.Base
{
	/// <summary>
	/// 
	/// </summary>
	public interface IElementServices
	{
		ElementEntity GetElementById(int ElementId);
		IEnumerable<ElementEntity> GetAllElements();
		int CreateElement(ElementEntity Element);
		bool UpdateElement(int ElementId, ElementEntity ElementEntity);
		bool DeleteElement(int ElementId);
	}
}
