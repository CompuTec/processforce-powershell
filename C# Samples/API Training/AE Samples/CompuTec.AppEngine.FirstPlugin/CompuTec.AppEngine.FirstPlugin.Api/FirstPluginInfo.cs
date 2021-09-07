using CompuTec.AppEngine.FirstPlugin.API.BusinessObjects.ToDo;
using CompuTec.Core2;
using System.Collections.Generic;
namespace CompuTec.AppEngine.FirstPlugin.API {
	

	/// <summary>
	/// This is a CoreInfo Dependency loader.
	/// this class is responsible to register all custom UDO Object imlementation 
	/// </summary>
	public class FirstPluginInfo : CoreInfo
	{
		public const string Name = "FirstPlugin";
		public const string NameVersion = "1.0.0.1";
		public const double DbVersion = 1.1d;
		//List of Busienss Object types that are implemented in this library
		private readonly List<string> implementedObjects = new List<string>();

		public FirstPluginInfo() : base(Name, NameVersion, DbVersion)
        {
			///Register UDO Object for Core2 repository 
            implementedObjects.Add(BusinessObjects.BusinessObjects.ToDoObjectCode);

        }

		/// <summary>
		/// Factory method that is used to create BusinessObjects based on type
		/// </summary>
		/// <param name="Token"></param>
		/// <param name="ObjectType"></param>
		/// <returns>returns null if this library does not implement this object 
		/// returns new instance of object for specified type
		/// </returns>
        public override dynamic CreateObject(string Token, string ObjectType)
		{
		
			if (ObjectType.Equals(BusinessObjects.BusinessObjects.ToDoObjectCode))
			{
				IToDo x = CoreManager.GetUDO<ToDo>(Token) as IToDo;
				return x;
			}
			return null;
		}
		/// <summary>
		/// This Function is used to determine the current instation number is used in update to obain if setup needs to be triggered
		/// </summary>
		/// <param name="Token"></param>
		/// <returns></returns>
		public override double GetCurrentDBVersion(string Token)
		{
			return 1.0d;
		}

		/// <summary>
		/// Inidicates weather this library implements specific businessObject
		/// </summary>
		/// <param name="ObjectType"></param>
		/// <returns>true if implements</returns>
		public override bool ImplementObject(string ObjectType)
		{
			bool implemented = implementedObjects.Contains(ObjectType);
			return implemented;
		}
	}
}