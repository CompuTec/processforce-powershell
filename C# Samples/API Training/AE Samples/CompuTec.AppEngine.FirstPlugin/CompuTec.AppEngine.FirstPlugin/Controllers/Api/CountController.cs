using CompuTec.AppEngine.Base.Infrastructure.Controllers.API;
using System;
using CompuTec.AppEngine.Base.Infrastructure.Plugins;
using System.Web.Http;

namespace CompuTec.AppEngine.FirstPlugin.Controllers.Api
{
    public class CountController : AppEngineController
    {
        [HttpGet]
        public int CountSupplierDocuments(string supplier)
        {
            var conf = Container.GetInstance<IPluginConfiguration>();

            var message = conf.Get<string>("Message");
            Console.WriteLine(message);

            int count = supplier.Length;
            return count;

        }

    }
}
