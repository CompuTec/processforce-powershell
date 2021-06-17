using CompuTec.AppEngine.DataAnnotations;
using CompuTec.Core2.Beans;
using CompuTec.Core2.Enumerators;
using System;

namespace CompuTec.AppEngine.First.Objects
{
	[AppEngineUDOBean(Ignore = false, ObjectType = "Sample_ToDo", TableName = "@SAMPLE_TODO")]
	public interface IToDo : IUDOBean
	{
		[AppEngineProperty(IsMasterKey = true)]
		String Code { get; set; }
		String Name { get; set; }
		DateTime UpdateDate { get; set; }
		DateTime U_Deadline { get; set; }
		string U_TaskName { get; set; }
		string U_Description { get; set; }
		string U_Priority { get; set; }

		IRequirement Requirements { get; set; }

	}
}
