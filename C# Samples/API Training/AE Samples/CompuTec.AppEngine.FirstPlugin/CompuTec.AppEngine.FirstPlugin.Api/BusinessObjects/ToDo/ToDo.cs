using CompuTec.Core2.Beans;
using System;
using System.Collections.Generic;

namespace CompuTec.AppEngine.FirstPlugin.API.BusinessObjects.ToDo
{
	public partial class ToDo
	{
		public ToDo()
		{
			this.UDOCode = BusinessObjects.ToDoObjectCode;
			this.TableName = "SAMPLE_OTDO";

			this.Childs = new Dictionary<string, ChildBeans>();
			this.ChildDictionary = new Dictionary<string, string>();

			this.Childs.Add("Requirements", new ToDoRequirement(true, this));
			this.ChildDictionary.Add("SAMPLE_TDO1", "Requirements");

		}

		protected override bool BeforeAdd()
		{

			this.U_Deadline = DateTime.Today.AddDays(7);
			this.Code = Guid.NewGuid().ToString();
			return base.BeforeAdd();
		}

		protected override bool BeforeUpdate()
		{

			return base.BeforeUpdate();
		}

	}
}