
using AzureAPI.ActionFilters;
using AzureAPI.ErrorHelper;
using BusinessEntities;
using BusinessServices;
using BusinessServices.Base;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace AzureAPI.Controllers
{
	/// <summary>
	/// Element controller class. you can add different classes like this. Token based security.
	/// </summary>
	[AuthorizationRequired]
	public class ElementController : ApiController
	{
		private readonly IElementServices _elementServices;

		#region Public Constructor

		/// <summary>
		/// Public constructor to initialize element service instance with DI
		/// 
		/// </summary>
		public ElementController(IElementServices elementServices)
		{
			_elementServices = elementServices;
		}

		#endregion

		// GET api/element
		/// <summary>
		/// Get the fullList of elements
		/// </summary>
		/// <returns>List of Elements</returns>
		public HttpResponseMessage Get()
		{
			try
			{
				var elements = _elementServices.GetAllElements();
				if (elements != null)
				{
					var elementEntities = elements as List<ElementEntity> ?? elements.ToList();
					if (elementEntities.Any())
						return Request.CreateResponse(HttpStatusCode.OK, elementEntities);

				}
				throw new ApiDataException(1001, "No product found", HttpStatusCode.NotFound);
			}
			catch
			{
				throw new ApiException() { ErrorCode = (int)HttpStatusCode.BadRequest, ErrorDescription = "Bad Request..." };
			}
		}

		// GET api/element/5
		/// <summary>
		/// Get a Prticular element
		/// </summary>
		/// <param name="id">Element Id</param>
		/// <returns>Element</returns>
		public HttpResponseMessage Get(int id)
		{
			if (id > 0)
			{
				var element = _elementServices.GetElementById(id);
				if (element != null)
				{
					return Request.CreateResponse(HttpStatusCode.OK, element);
				}
				throw new ApiDataException(1001, "No product found for this id.", HttpStatusCode.NotFound);
			}
			throw new ApiException() { ErrorCode = (int)HttpStatusCode.BadRequest, ErrorDescription = "Bad Request..." };
		}

		// POST api/element
		/// <summary>
		/// Creates a new element
		/// </summary>
		/// <param name="elementEntity">Parameters set</param>
		/// <returns>Id of the element</returns>
		public int Post([FromBody] ElementEntity elementEntity)
		{
			if (elementEntity != null)
			{
				return _elementServices.CreateElement(elementEntity);
			}
			throw new ApiBusinessException(2002, "Bad input.", HttpStatusCode.BadRequest);
		}

		// PUT api/element/5
		/// <summary>
		/// Update an element
		/// </summary>
		/// <param name="id">Element Id</param>
		/// <param name="elementEntity">SParameters set</param>
		/// <returns></returns>
		public bool Put(int id, [FromBody]ElementEntity elementEntity)
		{
			if (id > 0)
			{
				return _elementServices.UpdateElement(id, elementEntity);
			}
			return false;
		}

		// DELETE api/element/5
		/// <summary>
		/// Deletes and element
		/// </summary>
		/// <param name="id">Element Id</param>
		/// <returns></returns>
		public bool Delete(int id)
		{
			if (id > 0)
				return _elementServices.DeleteElement(id);
			return false;
		}
	}
}
