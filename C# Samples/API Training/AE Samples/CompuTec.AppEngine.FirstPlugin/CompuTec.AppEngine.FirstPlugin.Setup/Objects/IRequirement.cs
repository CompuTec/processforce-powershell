using CompuTec.AppEngine.DataAnnotations;
using CompuTec.Core2.Beans;
using CompuTec.Core2.Enumerators;
using System;
using System.Collections.Generic;

namespace CompuTec.AppEngine.First.Objects
{
	[AppEngineUDOChildBean()]
	public interface IRequirement : IUDOChildBean, IEnumerable<IRequirement>
	{
		String U_Name { get; set; }
        int U_Quantity { get; set; }
        new int U_LineNum { get; set; }

    }
}
