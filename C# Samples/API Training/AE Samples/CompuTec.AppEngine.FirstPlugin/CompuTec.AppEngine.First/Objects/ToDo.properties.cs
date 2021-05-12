using CompuTec.AppEngine.First.Objects;
using CompuTec.Core2.Beans;
using CompuTec.Core2.DI.Attributes;
using CompuTec.Core2.Enumerators;
using System;

namespace CompuTec.AppEngine.First.Objects
{
	public partial class ToDo : UDOBean, IToDo
	{
		public String Code
		{
			get { return FieldDictionary["Code"].Value; }
			set { FieldDictionary["Code"].Value = value; }
		}
		public String Name
		{
			get { return FieldDictionary["Name"].Value; }
			set { FieldDictionary["Name"].Value = value; ; }
		}
		public DateTime U_Deadline
		{
			get { return FieldDictionary["U_Deadline"].Value; }
			set { FieldDictionary["U_Deadline"].Value = value; ; }
		}




	}
}