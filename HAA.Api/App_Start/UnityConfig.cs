using HAA.Resolver;
using Microsoft.Practices.Unity;
using System.Web.Http;
using Unity.WebApi;

namespace HAA.Api
{
    public static class UnityConfig
	{
		public static void RegisterComponents()
		{
			var container = BuildUnityContainer();

			// register all your components with the container here
			// it is NOT necessary to register your controllers
			// e.g. container.RegisterType<ITestService, TestService>();
            //DependencyResolver.SetResolver(new UnityDependencyResolver(container));

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
			ComponentLoader.LoadContainer(container, ".\\bin", "HAA.Api.dll");
			ComponentLoader.LoadContainer(container, ".\\bin", "HAA.BusinessServices.dll");
            ComponentLoader.LoadContainer(container, ".\\bin", "HAA.DataModel.dll");
        }
	}
}