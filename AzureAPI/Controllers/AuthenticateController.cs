using AzureAPI.ErrorHelper;
using AzureAPI.Filters;
using BusinessServices.Base;
using Microsoft.Owin.Security.Cookies;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace AzureAPI.Controllers
{
	public class AuthenticateController : ApiController
	{
		#region Private variable.

		private readonly ITokenServices _tokenServices;

		#endregion

		#region Public Constructor

		/// <summary>
		/// Public constructor to initialize product service instance
		/// </summary>
		public AuthenticateController(ITokenServices tokenServices)
		{
			_tokenServices = tokenServices;
		}

		#endregion

		/// <summary>
		/// Test function will be a simple function with few variable and will return "Ok" output with the access token is valid. A test function will be developed to accept the access token and input and return valid output. Test function will be a simple function with few variable and will return "Ok" output with the access token is valid. There will not be any calculation inside the function.
		/// </summary>
		/// <returns></returns>
		[Route("TestToken")]
		[ApiAuthenticationFilter(false)]
		public HttpResponseMessage TestToken()
		{
			//If the funtion is called, the API has already validated the token
			return Request.CreateResponse(HttpStatusCode.OK, "OK");
		}

		/// <summary>
		/// Authenticates user and returns token with expiry. Set authenticatoin header ( Basic  encode64(user:password)). 
		/// You can convert your user:password to base64 at https://www.base64encode.org/
		/// </summary>
		/// <returns></returns>
		[Route("Login")]
		[ApiAuthenticationFilter]
		public HttpResponseMessage Login()
		{
			if (System.Threading.Thread.CurrentPrincipal != null && System.Threading.Thread.CurrentPrincipal.Identity.IsAuthenticated)
			{
				var basicAuthenticationIdentity = System.Threading.Thread.CurrentPrincipal.Identity as BasicAuthenticationIdentity;
				if (basicAuthenticationIdentity != null)
				{
					var userId = basicAuthenticationIdentity.UserId;
					return GetAuthToken(userId);
				}
			}
			throw new ApiException() { ErrorCode = (int)HttpStatusCode.BadRequest, ErrorDescription = "Bad Request..." };
		}

		/// <summary>
		/// Method for Logging out
		/// </summary>
		/// <returns></returns>
		[Route("Logout")]
		[ApiAuthenticationFilter(false)]
		[HttpPost]
		public HttpResponseMessage Logout()
		{
			var Token = "Token";

			if (Request.Headers.Contains(Token))
			{
				var tokenValue = Request.Headers.GetValues(Token);
				return KillAuthToken(Token);
			}
			throw new ApiException() { ErrorCode = (int)HttpStatusCode.BadRequest, ErrorDescription = "Bad Request..." };
		}
		
		/// <summary>
		/// Returns auth token for the validated user.
		/// </summary>
		/// <param name="userId"></param>
		/// <returns></returns>
		private HttpResponseMessage GetAuthToken(int userId)
		{
			var token = _tokenServices.GenerateToken(userId);
			var response = Request.CreateResponse(HttpStatusCode.OK, "Authorized");
			response.Headers.Add("Token", token.AuthToken);
			response.Headers.Add("TokenExpiry", ConfigurationManager.AppSettings["AuthTokenExpiry"]);
			response.Headers.Add("Access-Control-Expose-Headers", "Token,TokenExpiry");
			return response;
		}

		/// <summary>
		/// Returns auth token for the validated user.
		/// </summary>
		/// <param name="Token"></param>
		/// <returns></returns>
		private HttpResponseMessage KillAuthToken(string Token)
		{
			var token = _tokenServices.Kill(Token);
			var response = Request.CreateResponse(HttpStatusCode.OK, "Logout");
			return response;
		}
	}
}
