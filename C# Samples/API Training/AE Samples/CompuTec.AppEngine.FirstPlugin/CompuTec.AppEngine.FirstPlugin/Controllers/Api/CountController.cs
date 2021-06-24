using CompuTec.AppEngine.Base.Infrastructure.Controllers.API;
using CompuTec.AppEngine.Base.Infrastructure.Controllers;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Http;

namespace CompuTec.AppEngine.FirstPlugin.Controllers.Api
{
    public class CountController : AppEngineController
    {
        [HttpGet]
        public int CountSupplierDocuments(string supplier)
        {
            int count = supplier.Length;
            return count;

        }

    }
}
