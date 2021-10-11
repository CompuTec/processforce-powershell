using Enums;
using System.Collections.Generic;
using System.Linq;
using CompuTec.AppEngine.Base.Infrastructure.Controllers.Serialization;
using CompuTec.AppEngine.Base.Infrastructure.Exceptions;
using CompuTec.AppEngine.Base.Infrastructure.Services;

namespace CompuTec.AppEngine.FirstPlugin.Serializer.Serializers.ToDo
{
    public partial class ToDoRequirementSerializer : UdoChildBeanSerializer<CompuTec.AppEngine.FirstPlugin.Models.Models.ToDo.ToDoRequirement, CompuTec.AppEngine.FirstPlugin.API.BusinessObjects.ToDo.IToDoRequirement>
    {
        public override CompuTec.AppEngine.FirstPlugin.Models.Models.ToDo.ToDoRequirement ToModel(CompuTec.AppEngine.FirstPlugin.API.BusinessObjects.ToDo.IToDoRequirement udoChild)
        {
            var udo = (udoChild as CompuTec.Core2.Beans.IAdvancedUDOChildBean).Parent as CompuTec.AppEngine.FirstPlugin.API.BusinessObjects.ToDo.IToDo;
            var model = new CompuTec.AppEngine.FirstPlugin.Models.Models.ToDo.ToDoRequirement();
            model.U_Name = udoChild.U_Name;
            model.U_Quantity = udoChild.U_Quantity;
            model.U_LineNum = udoChild.U_LineNum;
            UDFsToModel(udoChild, model);
            return model;
        }

        public override CompuTec.AppEngine.FirstPlugin.API.BusinessObjects.ToDo.IToDoRequirement Update(CompuTec.AppEngine.FirstPlugin.API.BusinessObjects.ToDo.IToDoRequirement udoChild, CompuTec.AppEngine.FirstPlugin.Models.Models.ToDo.ToDoRequirement model)
        {
            var udo = (udoChild as CompuTec.Core2.Beans.IAdvancedUDOChildBean).Parent as CompuTec.AppEngine.FirstPlugin.API.BusinessObjects.ToDo.IToDo;
            if (model.U_Name != null)
                udoChild.U_Name = model.U_Name;
            if (model.U_Quantity != null)
            {
                udoChild.U_Quantity = (int)model.U_Quantity;
            }
            else
            {
                udoChild.U_Quantity = default(int);
            }

            if (model.U_LineNum != null)
            {
                udoChild.U_LineNum = (int)model.U_LineNum;
            }
            else
            {
                udoChild.U_LineNum = default(int);
            }

            UDFsToUdo(udoChild, model);
            return udoChild;
        }

        public override CompuTec.AppEngine.FirstPlugin.API.BusinessObjects.ToDo.IToDoRequirement FillNew(CompuTec.AppEngine.FirstPlugin.API.BusinessObjects.ToDo.IToDoRequirement udoChild, CompuTec.AppEngine.FirstPlugin.Models.Models.ToDo.ToDoRequirement model)
        {
            var udo = (udoChild as CompuTec.Core2.Beans.IAdvancedUDOChildBean).Parent as CompuTec.AppEngine.FirstPlugin.API.BusinessObjects.ToDo.IToDo;
            if (model.U_Name != null)
                udoChild.U_Name = model.U_Name;
            if (model.U_Quantity != null)
                udoChild.U_Quantity = (int)model.U_Quantity;
            if (model.U_LineNum != null)
                udoChild.U_LineNum = (int)model.U_LineNum;
            UDFsToUdo(udoChild, model);
            return udoChild;
        }

        protected override CompuTec.AppEngine.FirstPlugin.API.BusinessObjects.ToDo.IToDoRequirement FillNewExtended(CompuTec.AppEngine.FirstPlugin.API.BusinessObjects.ToDo.IToDoRequirement udoChild, CompuTec.AppEngine.FirstPlugin.Models.Models.ToDo.ToDoRequirement model)
        {
            var udo = (udoChild as CompuTec.Core2.Beans.IAdvancedUDOChildBean).Parent as CompuTec.AppEngine.FirstPlugin.API.BusinessObjects.ToDo.IToDo;
            if (model.U_Name != null)
                udoChild.U_Name = model.U_Name;
            if (model.U_Quantity != null)
                udoChild.U_Quantity = (int)model.U_Quantity;
            if (model.U_LineNum != null)
                udoChild.U_LineNum = (int)model.U_LineNum;
            UDFsToUdo(udoChild, model);
            return udoChild;
        }
    }
}