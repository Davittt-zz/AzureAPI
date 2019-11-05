using HAA.BusinessServices.Base;
using HAA.Resolver;
using System.ComponentModel.Composition;

namespace HAA.BusinessServices
{
    [Export(typeof(IComponent))]
	public class DependencyResolver : IComponent
	{
		public void SetUp(IRegisterComponent registerComponent)
		{
			registerComponent.RegisterType<IElementServices, ElementServices>();
			registerComponent.RegisterType<IUserServices, UserServices>();
			registerComponent.RegisterType<ITokenServices, TokenServices>();
            registerComponent.RegisterType<ISpeakerService, SpeakerService>();
        }
	}
}