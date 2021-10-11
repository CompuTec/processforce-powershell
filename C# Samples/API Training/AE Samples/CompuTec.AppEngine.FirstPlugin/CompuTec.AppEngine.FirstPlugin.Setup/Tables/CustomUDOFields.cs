using CompuTec.Core2.DI.Setup.UDO.Model;
using System.Collections.Generic;

namespace CompuTec.AppEngine.FirstPlugin.Setup.Tables
{
	public class CustomUDOFields
	{

		/// <summary>
		/// in this method you can list all UDF fields that needs to be installed with the plugin 
		/// </summary>
		/// <returns></returns>
		public static List<ICustomField> getCustomFields()
		{
			var list = new List<ICustomField>();
			///UDF U_FirstUDF of type alphanumeric 20 added to Items object
			var udf = new TableCustomField();
			udf.SetName("FistUDF");
			udf.SetDescription("Fisrst Plugin First UDF");
			udf.SetTableName("OITM");
			udf.SetEditSize(20);
			udf.SetType(BaseLayer.DI.BoFieldTypes.db_Alpha);
			list.Add(udf);

			return list;
		}
	}
}
