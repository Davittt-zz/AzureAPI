using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Http;
using System.Web.Mvc;
using System.Web.Optimization;
using System.Web.Routing;

namespace AzureAPI
{
	public class WebApiApplication : System.Web.HttpApplication
	{
		public object Bootstrapper { get; private set; }

		protected void Application_Start()
		{
			AreaRegistration.RegisterAllAreas();
			//Initialise Unity register
			UnityConfig.RegisterComponents();

			GlobalConfiguration.Configure(WebApiConfig.Register);
			FilterConfig.RegisterGlobalFilters(GlobalFilters.Filters);
			RouteConfig.RegisterRoutes(RouteTable.Routes);
			BundleConfig.RegisterBundles(BundleTable.Bundles);

		

	//	//Define Formatters
	//	var formatters = GlobalConfiguration.Configuration.Formatters;
	//	var jsonFormatter = formatters.JsonFormatter;
	//	var settings = jsonFormatter.SerializerSettings;
	//	settings.Formatting = Formatting.Indented;
	//	// settings.ContractResolver = new CamelCasePropertyNamesContractResolver();
	//	var appXmlType = formatters.XmlFormatter.SupportedMediaTypes.FirstOrDefault(t => t.MediaType == "application/xml");
	//	formatters.XmlFormatter.SupportedMediaTypes.Remove(appXmlType);
	//
	//	//Add CORS Handler
	//	GlobalConfiguration.Configuration.MessageHandlers.Add(new CorsHandler());
		}
	}
}
