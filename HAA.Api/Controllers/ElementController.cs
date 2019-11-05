
using HAA.Api.ActionFilters;
using HAA.Api.ErrorHelper;
using HAA.BusinessEntities;
using HAA.BusinessServices.Base;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace HAA.Api.Controllers
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
		/// Get the fullList of elements. Add Header with Token "authToken". you can find it in your browser (Chrome->PressF12->Application Tab->)
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
				return Request.CreateResponse(HttpStatusCode.OK, "");
				//throw new ApiDataException(1001, "No product found", HttpStatusCode.NotFound);
			}
			catch 
			{
				throw new ApiException() { ErrorCode = (int)HttpStatusCode.BadRequest, ErrorDescription = "Bad Request..." };
			}
		}

		// GET api/element/5
		/// <summary>
		/// Get a Particular element. Add Header with Token "authToken"
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
		/// Creates a new element. Add Header with Token "authToken"
		/// </summary>
		/// <param name="elementEntity">Parameters set</param>
		/// <returns>Id of the element</returns>
		public HttpResponseMessage Post([FromBody] ElementEntity elementEntity)
		{
			try
			{
				if (elementEntity != null)
				{
					if (ModelState.IsValid)
					{
						return Request.CreateResponse(HttpStatusCode.OK, _elementServices.CreateElement(elementEntity));
					}
					else
					{
						throw new ApiBusinessException(2003, "Bad Request. Invalid object", HttpStatusCode.BadRequest);
					}
				}
				throw new ApiBusinessException(2002, "Bad Request. Null element", HttpStatusCode.BadRequest);
			}
			catch
			{
				throw new ApiDataException(3001, "Internal error", HttpStatusCode.InternalServerError);
			}
		}

		// PUT api/element/5
		/// <summary>
		/// Update an element. Add Header with Token "authToken"
		/// </summary>
		/// <param name="id">Element Id</param>
		/// <param name="elementEntity">SParameters set</param>
		/// <returns></returns>
		public HttpResponseMessage Put(int id, [FromBody]ElementEntity elementEntity)
		{
			try
			{
				if (elementEntity != null)
				{
					if (ModelState.IsValid)
					{
						if (id > 0)
						{
							return Request.CreateResponse(HttpStatusCode.OK, _elementServices.UpdateElement(id, elementEntity));
						}
						else {
							throw new ApiBusinessException(2004, "Object not updated. There is no Element with Id:" + id, HttpStatusCode.NotModified);
						}
					}
					else
					{
						throw new ApiBusinessException(2003, "Bad Request. Invalid object", HttpStatusCode.BadRequest);
					}
				}
				throw new ApiBusinessException(2002, "Bad Request. Null element", HttpStatusCode.BadRequest);
			}
			catch
			{
				throw new ApiDataException(3001, "Internal error", HttpStatusCode.InternalServerError);
			}
		}
		// DELETE api/element/5
		/// <summary>
		/// Deletes and element. Add Header with Token "authToken"
		/// </summary>
		/// <param name="id">Element Id</param>
		/// <returns></returns>
		public HttpResponseMessage Delete(int id)
		{
			if (id > 0)
			{
				return Request.CreateResponse(HttpStatusCode.OK, _elementServices.DeleteElement(id));
			}
			else {
				throw new ApiBusinessException(2004, "Object not updated. There is no Element with Id:" + id, HttpStatusCode.NotModified);
			}
		}
	}
}
