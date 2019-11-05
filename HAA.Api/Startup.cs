using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Owin;
using Owin;

[assembly: OwinStartup(typeof(HAA.Api.Startup))]

namespace HAA.Api
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
      //      ConfigureAuth(app);
        }
    }
}
