using CompuTec.AppEngine.DataAnnotations;
using CompuTec.AppEngine.FirstPlugin.API.Enums;
using CompuTec.Core2.Beans;
using CompuTec.Core2.Enumerators;
using System;
using System.ComponentModel;

namespace CompuTec.AppEngine.FirstPlugin.API.BusinessObjects.ToDo
{


	/// <summary>
	/// public interface that is exposed to 3rd party application - can be used in powershell import etc
	/// 
	/// AppEngine Annonations are used to descripte REST and OData Modes and Serializers used in Plugin controlers
	/// </summary>
	[AppEngineUDOBean(Ignore = false, ObjectType = "Sample_ToDo", TableName = "@CT_TST_OTDO")]
	public interface IToDo : IUDOBean
	{
		[AppEngineProperty(IsMasterKey = true)]
		String Code { get; set; }
		String Name { get; set; }
		DateTime UpdateDate { get; set; }
		DateTime U_Deadline { get; set; }
		string U_TaskName { get; set; }
		string U_Description { get; set; }
		[DefaultValue(ToDoPriority.Medium)]
		ToDoPriority U_Priority { get; set; }
		IToDoRequirement Requirements { get; set; }

	}
}
