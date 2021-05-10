using CompuTec.AppEngine.DataAnnotations;
using CompuTec.Core2.Beans;
using CompuTec.Core2.Enumerators;
using System;

namespace CompuTec.AppEngine.First.Objects
{
	[AppEngineUDOBean(Ignore = false, ObjectType = "Sample_ToDo", TableName = "@Sample_OTDO")]
	public interface IToDo : IUDOBean
	{
		[AppEngineProperty(IsMasterKey = true)]
		String Code { get; set; }
		String Name { get; set; }
        DateTime Deadline { get; set; }

    }
}
