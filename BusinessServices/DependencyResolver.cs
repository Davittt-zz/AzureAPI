using BusinessServices.Base;
using Resolver;
using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessServices
{
	[Export(typeof(IComponent))]
	public class DependencyResolver : IComponent
	{
		public void SetUp(IRegisterComponent registerComponent)
		{
			registerComponent.RegisterType<IElementServices, ElementServices>();
			registerComponent.RegisterType<IUserServices, UserServices>();
			registerComponent.RegisterType<ITokenServices, TokenServices>();
		}
	}
}
