using CompuTec.AppEngine.DataAnnotations;
using CompuTec.AppEngine.FirstPlugin.API.Enums;
using CompuTec.Core2.Beans;
using System;
using System.ComponentModel;

namespace CompuTec.AppEngine.FirstPlugin.API.BusinessObjects.ToDo
{
	/// <summary>
	/// Public interface that is exposed to 3rd party application - can be used in powershell import etc
	/// 
	/// AppEngine Annonations are used to descripte REST and OData Modes and Serializers used in Plugin controlers
	/// </summary>
	[AppEngineUDOBean(Ignore = false, ObjectType = "SAMPLE_TO_DO", TableName = "@SAMPLE_OTDO")]
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
		[DefaultValue(YesNoType.No)]
		YesNoType U_Done { get; set; }
		IToDoRequirement Requirements { get; set; }
	}
}
