using CompuTec.Core2.Beans;
using CompuTec.Core2.Enumerators;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CompuTec.AppEngine.First.Objects
{
	public partial class ToDo
	{
		

		public ToDo()
		{
            this.UDOCode = First.DBInstall.Tables.ToDoTable.OBJECT_CODE;
            this.TableName = First.DBInstall.Tables.ToDoTable.TABLE_NAME;

			this.Childs = new Dictionary<string, ChildBeans>();
			this.ChildDictionary = new Dictionary<string, string>();

			this.Childs.Add("Requirements", new Requirement());
			this.ChildDictionary.Add(DBInstall.Tables.RequirementTable.TABLE_NAME, "Requirements");
		}

	

		protected override bool BeforeAdd()
		{
			this.U_Deadline = DateTime.Now; 
			//this.Code = "Sample_Code";
			
			return base.BeforeAdd();
		}

		protected override bool BeforeUpdate()
		{
			
			return base.BeforeUpdate();
		}

		

		
	}
}