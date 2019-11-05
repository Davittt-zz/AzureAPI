using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace HAA.Api.Controllers
{
	/// <summary>
	/// Value class is a set of metods which are generated automatically. The class has no security mechanisms
	/// </summary>
	public class ValuesController : ApiController
	{
		// GET api/values
		/// <summary>
		/// Get the fullList of Values
		/// </summary>
		/// <returns></returns>
		public IEnumerable<string> Get()
		{
			return new string[] { "value1", "value2" };
		}

		// GET api/values/5
		/// <summary>
		/// Get a Prticular Value
		/// </summary>
		/// <param name="id">Value Id</param>
		/// <returns></returns>
		public string Get(int id)
		{
			return "value";
		}

		// POST api/values
		/// <summary>
		/// Creates a new Value
		/// </summary>
		/// <param name="value">Parameters set</param>
		/// <returns>Id of the element</returns>
		public void Post([FromBody]string value)
		{
		}

		// PUT api/values/5
		/// <summary>
		/// Edit Value
		/// </summary>
		/// <param name="id">Value Id</param>
		/// <param name="value">Parameters Set</param>
		public void Put(int id, [FromBody]string value)
		{
		}

		/// <summary>
		/// Deletes Value
		/// </summary>
		/// <param name="id">Value Id</param>
		public void Delete(int id)
		{
		}
	}
}
