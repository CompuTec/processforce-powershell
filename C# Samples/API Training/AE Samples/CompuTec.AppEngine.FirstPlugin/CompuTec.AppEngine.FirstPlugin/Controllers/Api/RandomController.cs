using CompuTec.AppEngine.Base.Infrastructure.Controllers.API;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Http;

namespace CompuTec.AppEngine.FirstPlugin.Controllers.Api
{
    public class RandomController : AppEngineController
    {
        public int Randomss()
        {
            int c = 2 + 2;
            return c;
        }
    }
}
