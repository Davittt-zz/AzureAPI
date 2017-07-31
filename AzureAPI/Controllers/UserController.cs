using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Security.Claims;
using System.Security.Cryptography;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;
using System.Web.Http.ModelBinding;
using Microsoft.AspNet.Identity;
using Microsoft.AspNet.Identity.EntityFramework;
using Microsoft.AspNet.Identity.Owin;
using Microsoft.Owin.Security;
using Microsoft.Owin.Security.Cookies;
using Microsoft.Owin.Security.OAuth;
using AzureAPI.Models;
using AzureAPI.Providers;
using AzureAPI.Results;
using AzureAPI.ActionFilters;
using BusinessEntities.Account;
using BusinessServices.Base;
using System.Net;
using AzureAPI.ErrorHelper;
using System.Threading;
using AzureAPI.Filters;

namespace AzureAPI.Controllers
{
	[AuthorizationRequired]
	public class UserController : ApiController
	{
		private readonly IUserServices _userServices;

		#region Public Constructor

		/// <summary>
		/// Public constructor to initialize element service instance with DI
		/// 
		/// </summary>
		public UserController(IUserServices userServices)
		{
			_userServices = userServices;
		}

		#endregion

		// GET api/User/UserInfo
		/// <summary>
		/// Get User information. Add Header with Token "authToken"
		/// </summary>
		/// <param name="id"></param>
		/// <returns></returns>
		[Route("UserInfo")]
		public HttpResponseMessage GetUserInfo()
		{
			var basicAuthenticationIdentity = Thread.CurrentPrincipal.Identity as BasicAuthenticationIdentity;
		//	if (basicAuthenticationIdentity != null)

				var id = basicAuthenticationIdentity.UserId;

			if (id > 0)
			{
				var user = _userServices.GetUserById(id);
				if (user != null)
				{
					return Request.CreateResponse(HttpStatusCode.OK, user);
				}
				throw new ApiDataException(1001, "No user found for this id.", HttpStatusCode.NotFound);
			}
			throw new ApiException() { ErrorCode = (int)HttpStatusCode.BadRequest, ErrorDescription = "Bad Request..." };
		}

		protected override void Dispose(bool disposing)
		{
			base.Dispose(disposing);
		}
	}
}
