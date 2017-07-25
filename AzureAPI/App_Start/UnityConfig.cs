using BusinessServices;
using BusinessServices.Base;
using DataModel.UnitOfWork;
using Microsoft.Practices.Unity;
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

		//	DependencyResolver.SetResolver(new UnityDependencyResolver(container));

			GlobalConfiguration.Configuration.DependencyResolver = new UnityDependencyResolver(container);
        }

		private static IUnityContainer BuildUnityContainer()
		{
			var container = new UnityContainer();
			// register all your components with the container here
			// it is NOT necessary to register your controllers
			container.RegisterType<IElementServices, ElementServices>().RegisterType<UnitOfWork>(new HierarchicalLifetimeManager());
			return container;
		}
	}
}