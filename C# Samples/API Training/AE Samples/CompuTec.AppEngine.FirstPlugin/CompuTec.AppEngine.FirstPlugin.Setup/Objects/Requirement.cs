using CompuTec.Core2.Beans;
using System;
using System.Collections.Generic;
using System.Linq;
using CompuTec.AppEngine.First.Objects;
using System.Text;
using System.Threading.Tasks;
using System.Collections;

namespace CompuTec.AppEngine.First.Objects
{
    public class Requirement : ChildBeans, IRequirement
    {
		public Requirement()
		{
		}
		public Requirement(ChildBeans childBeans) : base(childBeans) { }
		public Requirement(bool master, UDOBean baseUDO) : base(master, baseUDO) { }

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
		 public IEnumerator<IRequirement> GetEnumerator()
		{
			return new ChildBeansEnum<IRequirement>(this);
		}
		IEnumerator IEnumerable.GetEnumerator()
		{
			return (IEnumerator<IRequirement>)GetEnumerator();
		}

      
    }
}
