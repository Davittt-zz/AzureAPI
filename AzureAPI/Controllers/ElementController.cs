using AzureAPI.Filters;
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
	[ApiAuthenticationFilter]
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
		/// <returns></returns>
		public HttpResponseMessage Get()
		{
			var elements = _elementServices.GetAllElements();
			if (elements != null)
			{
				var elementEntities = elements as List<ElementEntity> ?? elements.ToList();
				if (elementEntities.Any())
					return Request.CreateResponse(HttpStatusCode.OK, elementEntities);
			}
			return Request.CreateErrorResponse(HttpStatusCode.NotFound, "Elements not found");
		}

		// GET api/element/5
		public HttpResponseMessage Get(int id)
		{
			var element = _elementServices.GetElementById(id);
			if (element != null)
				return Request.CreateResponse(HttpStatusCode.OK, element);
			return Request.CreateErrorResponse(HttpStatusCode.NotFound, "No element found for this id");
		}

		// POST api/element
		public int Post([FromBody] ElementEntity elementEntity)
		{
			return _elementServices.CreateElement(elementEntity);
		}

		// PUT api/element/5
		public bool Put(int id, [FromBody]ElementEntity elementEntity)
		{
			if (id > 0)
			{
				return _elementServices.UpdateElement(id, elementEntity);
			}
			return false;
		}

		// DELETE api/element/5
		public bool Delete(int id)
		{
			if (id > 0)
				return _elementServices.DeleteElement(id);
			return false;
		}
	}
}
