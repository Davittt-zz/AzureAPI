using System.Collections.Generic;
using BusinessEntities;

namespace BusinessServices.Base
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
