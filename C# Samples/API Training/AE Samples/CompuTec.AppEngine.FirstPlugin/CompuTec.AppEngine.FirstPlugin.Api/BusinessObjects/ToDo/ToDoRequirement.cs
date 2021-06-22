using CompuTec.Core2.Beans;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;

namespace CompuTec.AppEngine.FirstPlugin.API.BusinessObjects.ToDo
{
    public class ToDoRequirement : ChildBeans, IToDoRequirement
    {
		public ToDoRequirement()
		{
		}
		public ToDoRequirement(ChildBeans childBeans) : base(childBeans) { }
		public ToDoRequirement(bool master, UDOBean baseUDO) : base(master, baseUDO) { }

		public String U_Name
		{
			get { return FieldDictionary["U_Name"].Value; }
			set { FieldDictionary["U_Name"].Value = value; }
		}
		public int U_Quantity
		{
			get { return FieldDictionary["U_Quantity"].Value; }
			set { FieldDictionary["U_Quantity"].Value = value; }
		}
		 public IEnumerator<IToDoRequirement> GetEnumerator()
		{
			return new ChildBeansEnum<IToDoRequirement>(this);
		}
		IEnumerator IEnumerable.GetEnumerator()
		{
			return (IEnumerator<IToDoRequirement>)GetEnumerator();
		}

      
    }
}
