using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Web.Http;
using Microsoft.Owin.Security.OAuth;
using Newtonsoft.Json.Serialization;
using AzureAPI.ActionFilters;

namespace AzureAPI
{
	public static class WebApiConfig
	{
		public static void Register(HttpConfiguration config)
		{

			// Configuración y servicios de Web API
			// Configure Web API para usar solo la autenticación de token de portador.
		//	config.SuppressDefaultHostAuthentication();
		//	config.Filters.Add(new HostAuthenticationFilter(OAuthDefaults.AuthenticationType));
					
			//Config Logging filter
			config.Filters.Add(new LoggingFilterAttribute());

			config.Filters.Add(new GlobalExceptionAttribute());

			//Config GlobalException filter
			//config.Filters.Add(new GlobalExceptionAttribute());

			// Rutas de Web API
			config.MapHttpAttributeRoutes();

			config.Routes.MapHttpRoute(
				name: "DefaultApi",
				routeTemplate: "api/{controller}/{id}",
				defaults: new { id = RouteParameter.Optional }
			);
		}
	}
}
