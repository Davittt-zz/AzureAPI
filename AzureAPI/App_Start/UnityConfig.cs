using BusinessServices;
using BusinessServices.Base;
using Microsoft.Practices.Unity;
using Resolver;
using System.Web.Http;
using Unity.WebApi;

namespace AzureAPI
{
	public static class UnityConfig
	{
		public static void RegisterComponents()
		{
			var container = BuildUnityContainer();

			// register all your components with the container here
			// it is NOT necessary to register your controllers
			// e.g. container.RegisterType<ITestService, TestService>();
//			DependencyResolver.SetResolver(new UnityDependencyResolver(container));

			GlobalConfiguration.Configuration.DependencyResolver = new UnityDependencyResolver(container);
		}

		private static IUnityContainer BuildUnityContainer()
		{
			var container = new UnityContainer();

			// register all your components with the container here
			// it is NOT necessary to register your controllers

			// e.g. container.RegisterType<ITestService, TestService>();       
			// container.RegisterType<IProductServices, ProductServices>().RegisterType<UnitOfWork>(new HierarchicalLifetimeManager());

			RegisterTypes(container);

			return container;
		}

		public static void RegisterTypes(IUnityContainer container)
		{
			//Component initialization via MEF
			ComponentLoader.LoadContainer(container, ".\\bin", "AzureAPI.dll");
			ComponentLoader.LoadContainer(container, ".\\bin", "BusinessServices.dll");
		}
	}
}