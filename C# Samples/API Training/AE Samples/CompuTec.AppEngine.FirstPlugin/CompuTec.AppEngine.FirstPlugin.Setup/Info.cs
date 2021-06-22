using CompuTec.AppEngine.First.Objects;
using CompuTec.Core2;
using System.Collections.Generic;

namespace CompuTec.AppEngine.First { 
	public class Info : CoreInfo
	{
		public const string Name = "FirstPlugin";
		public const string NameVersion = "1.0.0.1";
		public const double DbVersion = 1.1d;

		private readonly List<string> implementedObjects = new List<string>();

		public Info() : base(Name, NameVersion, DbVersion)
		{
			implementedObjects.Add(DBInstall.Tables.ToDoTable.OBJECT_CODE);
		}
		public override dynamic CreateObject(string Token, string ObjectType)
		{
			if (ObjectType.Equals(DBInstall.Tables.ToDoTable.OBJECT_CODE))
			{

				IToDo x = CoreManager.GetUDO<ToDo>(Token) as IToDo;
				return x;
			}
			return null;
		}

		public override double GetCurrentDBVersion(string Token)
		{
			return 1.0d;
		}


		public override bool ImplementObject(string ObjectType)
		{
			bool implemented = implementedObjects.Contains(ObjectType);
			return implemented;
		}
	}
}