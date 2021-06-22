using CompuTec.Core2.Beans;
using CompuTec.Core2.Enumerators;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CompuTec.AppEngine.FirstPlugin.API.BusinessObjects.ToDo
{
	public partial class ToDo
	{
		

		public ToDo()
		{
            this.UDOCode = BusinessObjects.ToDoObjectCode;
            this.TableName = "CT_TST_OTDO";

			this.Childs = new Dictionary<string, ChildBeans>();
			this.ChildDictionary = new Dictionary<string, string>();

			this.Childs.Add("Requirements", new ToDoRequirement());
			this.ChildDictionary.Add("CT_TST_TDO1", "Requirements");

		}
		 
		protected override bool BeforeAdd()
		{
           
			this.U_Deadline = DateTime.Today.AddDays(7);
			this.Code = (this.U_TaskName + "Code").ToString();
			
			return base.BeforeAdd();
		}

		protected override bool BeforeUpdate()
		{	
			
			return base.BeforeUpdate();
		}

		

		
	}
}