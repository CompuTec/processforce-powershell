using CompuTec.AppEngine.FirstPlugin.API.Enums;
using CompuTec.Core2.Beans;
using System;

namespace CompuTec.AppEngine.FirstPlugin.API.BusinessObjects.ToDo
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
			set { FieldDictionary["Name"].Value = value; }
		}
		public DateTime UpdateDate
		{
			get { return FieldDictionary["UpdateDate"].Value; }
			set { FieldDictionary["UpdateDate"].Value = value; }
		}
		public DateTime U_Deadline
		{
			get { return FieldDictionary["U_Deadline"].Value; }
			set { FieldDictionary["U_Deadline"].Value = value; }
		}
		public string U_TaskName
		{
			get { return FieldDictionary["U_TaskName"].Value; }
			set { FieldDictionary["U_TaskName"].Value = value; }
		}
		public string U_Description
		{
			get { return FieldDictionary["U_Description"].Value; }
			set { FieldDictionary["U_Description"].Value = value; }
		}
		public ToDoPriority U_Priority
		{
			get { return FieldDictionary["U_Priority"].Value; }
			set { FieldDictionary["U_Priority"].Value = value; }
		}
		public YesNoType U_Done
		{
			get { return FieldDictionary["U_Done"].Value; }
			set { FieldDictionary["U_Done"].Value = value; }
		}
		public IToDoRequirement Requirements
		{
			get => this.Childs["Requirements"].CurrentChild as IToDoRequirement;
			set { this.Childs["Requirements"].CurrentChild = value as IToDoRequirement; }
		}
	
	}
}