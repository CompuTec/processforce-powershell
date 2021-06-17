using System.Collections.Generic;
using System.Linq;
using CompuTec.AppEngine.Base.Infrastructure.Controllers.Serialization;
using CompuTec.AppEngine.Base.Infrastructure.Exceptions;
using CompuTec.AppEngine.Base.Infrastructure.Services;

namespace CompuTec.AppEngine.FirstPlugin.Serializer.Serializers.Objects
{
    public partial class RequirementSerializer : UdoChildBeanSerializer<CompuTec.AppEngine.FirstPlugin.Models.Models.Objects.Requirement, CompuTec.AppEngine.First.Objects.IRequirement>
    {
        public override CompuTec.AppEngine.FirstPlugin.Models.Models.Objects.Requirement ToModel(CompuTec.AppEngine.First.Objects.IRequirement udoChild)
        {
            var udo = (udoChild as CompuTec.Core2.Beans.IAdvancedUDOChildBean).Parent as CompuTec.AppEngine.First.Objects.IToDo;
            var model = new CompuTec.AppEngine.FirstPlugin.Models.Models.Objects.Requirement();
            model.Code = udoChild.Code;
            model.Name = udoChild.Name;
            UDFsToModel(udoChild, model);
            return model;
        }

        public override CompuTec.AppEngine.First.Objects.IRequirement Update(CompuTec.AppEngine.First.Objects.IRequirement udoChild, CompuTec.AppEngine.FirstPlugin.Models.Models.Objects.Requirement model)
        {
            var udo = (udoChild as CompuTec.Core2.Beans.IAdvancedUDOChildBean).Parent as CompuTec.AppEngine.First.Objects.IToDo;
            if (model.Code != null)
                udoChild.Code = model.Code;
            if (model.Name != null)
                udoChild.Name = model.Name;
            UDFsToUdo(udoChild, model);
            return udoChild;
        }

        public override CompuTec.AppEngine.First.Objects.IRequirement FillNew(CompuTec.AppEngine.First.Objects.IRequirement udoChild, CompuTec.AppEngine.FirstPlugin.Models.Models.Objects.Requirement model)
        {
            var udo = (udoChild as CompuTec.Core2.Beans.IAdvancedUDOChildBean).Parent as CompuTec.AppEngine.First.Objects.IToDo;
            if (model.Code != null)
                udoChild.Code = model.Code;
            if (model.Name != null)
                udoChild.Name = model.Name;
            UDFsToUdo(udoChild, model);
            return udoChild;
        }

        protected override CompuTec.AppEngine.First.Objects.IRequirement FillNewExtended(CompuTec.AppEngine.First.Objects.IRequirement udoChild, CompuTec.AppEngine.FirstPlugin.Models.Models.Objects.Requirement model)
        {
            var udo = (udoChild as CompuTec.Core2.Beans.IAdvancedUDOChildBean).Parent as CompuTec.AppEngine.First.Objects.IToDo;
            if (model.Code != null)
                udoChild.Code = model.Code;
            if (model.Name != null)
                udoChild.Name = model.Name;
            UDFsToUdo(udoChild, model);
            return udoChild;
        }
    }
}