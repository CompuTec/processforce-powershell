using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CompuTec.AppEngine.FirstPlugin.API.Enums
{
    [CompuTec.Core2.DI.Attributes.EnumType(new int[] { 1, 2, 3 },new string[] { "L" ,"M","H"},2)]

    public enum  ToDoPriority
    {
        Low=1,Medium=2,Huge=3
    }
}
