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
		public int Quantity
		{
			get { return FieldDictionary["Quantity"].Value; }
			set { FieldDictionary["U_Quantity"].Value = value; }
		}
		 public IEnumerator<Requirement> GetEnumerator()
		{
			return new ChildBeansEnum<Requirement>(this);
		}
		IEnumerator IEnumerable.GetEnumerator()
		{
			return (IEnumerator<Requirement>)GetEnumerator();
		}

        IEnumerator<IRequirement> IEnumerable<IRequirement>.GetEnumerator()
        {
            throw new NotImplementedException();
        }
    }
}
