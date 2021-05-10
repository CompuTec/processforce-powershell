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
            //this.Childs = new Dictionary<string, Core2.Beans.ChildBeans>();
            this.UDOCode = First.DBInstall.Tables.ToDoTable.OBJECT_CODE;
            this.TableName = First.DBInstall.Tables.ToDoTable.TABLE_NAME;


        }

		protected override bool BeforeAdd()
		{
			this.Deadline = DateTime.Now;
			
			return base.BeforeAdd();
		}

		protected override bool BeforeUpdate()
		{
			
			
			return base.BeforeUpdate();
		}

		

		
	}
}