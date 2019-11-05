using HAA.Api.ActionFilters;
using System.Web.Http;

namespace HAA.Api
{
    public static class WebApiConfig
	{
		public static void Register(HttpConfiguration config)
		{

			// Configuración y servicios de Web API
			// Configure Web API para usar solo la autenticación de token de portador.
			//config.SuppressDefaultHostAuthentication();
			//config.Filters.Add(new HostAuthenticationFilter(OAuthDefaults.AuthenticationType));

			//Config Logging filter
			config.Filters.Add(new LoggingFilterAttribute());

			//Config GlobalException filter
			config.Filters.Add(new GlobalExceptionAttribute());

			// Web API Routes
			config.MapHttpAttributeRoutes();

			config.Routes.MapHttpRoute(
				name: "DefaultApi",
				routeTemplate: "api/{controller}/{id}",
				defaults: new { id = RouteParameter.Optional }
			);
		}
	}
}
