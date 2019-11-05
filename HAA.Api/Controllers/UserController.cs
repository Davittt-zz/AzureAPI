using HAA.Api.ActionFilters;
using HAA.Api.ErrorHelper;
using HAA.BusinessServices.Base;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace HAA.Api.Controllers
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
		public HttpResponseMessage GetUserInfo(int id)
		{
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
