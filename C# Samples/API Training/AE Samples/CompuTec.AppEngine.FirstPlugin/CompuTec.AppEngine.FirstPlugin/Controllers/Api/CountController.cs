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
            //string connectionString = ;
            SqlConnection conn = new SqlConnection("connectionString");
            conn.Open();
            SqlCommand comm = new SqlCommand($"SELECT COUNT(*) FROM ORDR WHERE CardName = '{supplier}'", conn);
            Int32 count = Convert.ToInt32(comm.ExecuteScalar());
            conn.Close();
            return count;

        }

    }
}
