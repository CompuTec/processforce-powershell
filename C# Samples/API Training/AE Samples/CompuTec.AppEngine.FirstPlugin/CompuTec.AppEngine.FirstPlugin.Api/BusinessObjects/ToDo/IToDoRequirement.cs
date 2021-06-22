using CompuTec.AppEngine.DataAnnotations;
using CompuTec.Core2.Beans;
using CompuTec.Core2.Enumerators;
using System;
using System.Collections.Generic;

namespace CompuTec.AppEngine.FirstPlugin.API.BusinessObjects.ToDo 
{
	[AppEngineUDOChildBean()]
	public interface IToDoRequirement : IUDOChildBean, IEnumerable<IToDoRequirement>
	{
		String U_Name { get; set; }
        int U_Quantity { get; set; }
        new int U_LineNum { get; set; }
    }
}
