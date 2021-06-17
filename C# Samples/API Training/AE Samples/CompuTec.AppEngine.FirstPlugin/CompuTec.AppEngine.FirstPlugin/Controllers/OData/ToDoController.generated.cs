using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Http;
using System.Web.Http.Description;
using System.Web.OData;
using System.Web.OData.Query;
using System.Web.OData.Routing;
using CompuTec.Core2.Beans;
using CompuTec.AppEngine.Base.Infrastructure.Annotation;
using CompuTec.AppEngine.Base.Infrastructure.Controllers;
using CompuTec.AppEngine.Base.Infrastructure.Controllers.OData;
using CompuTec.AppEngine.Base.Infrastructure.Controllers.OData.Delta;
using CompuTec.AppEngine.Base.Infrastructure.Controllers.Serialization;
using CompuTec.AppEngine.Base.Infrastructure.Exceptions;
using CompuTec.AppEngine.FirstPlugin.Serializer.Serializers;

namespace CompuTec.AppEngine.FirstPlugin.Controllers.OData
{
    [ODataRoutePrefix("ToDo")]
    public partial class ToDoController : AppEngineCore2ODataBatchController<CompuTec.AppEngine.FirstPlugin.Models.Models.Objects.ToDo, CompuTec.AppEngine.First.Objects.IToDo, string>
    {
        protected override string TableName => "@SAMPLE_TODO";
        protected override string KeyName => "Code";
        protected override string ObjectType => "Sample_ToDo";
        protected override eUDOVersion UDOVersion => eUDOVersion.UDO_20;
        [HttpGet]
        [EnableQuery(MaxExpansionDepth = 10)]
        [ODataRoute("({Code})/Requirements")]
        public IQueryable<CompuTec.AppEngine.FirstPlugin.Models.Models.Objects.Requirement> GetRequirement([FromODataUri] string Code)
        {
            var udo = GetObjectInstance(Code);
            var model = Serializer.ToModel(udo);
            var Requirements = model.Requirements;
            return Requirements.AsQueryable<CompuTec.AppEngine.FirstPlugin.Models.Models.Objects.Requirement>();
        }

        [HttpGet]
        [EnableQuery(MaxExpansionDepth = 10)]
        [ODataRoute("({Code})/Requirements({RequirementsLineNum})")]
        public SingleResult<CompuTec.AppEngine.FirstPlugin.Models.Models.Objects.Requirement> GetRequirement([FromODataUri] string Code, [FromODataUri] int RequirementsLineNum)
        {
            var udo = GetObjectInstance(Code);
            var model = Serializer.ToModel(udo);
            var Requirements = model.Requirements.FirstOrDefault(item => RequirementsLineNum == item.U_LineNum);
            if (Requirements == null)
                throw new NotFoundException("Requirements", "Requirements");
            return SingleResult.Create(new List<CompuTec.AppEngine.FirstPlugin.Models.Models.Objects.Requirement>()
            {Requirements}.AsQueryable());
        }

        [HttpPost]
        [ODataRoute("({Code})/Requirements")]
        public IHttpActionResult PostRequirement([FromODataUri] string Code, CompuTec.AppEngine.FirstPlugin.Models.Models.Objects.Requirement model)
        {
            var udo = GetObjectInstance(Code);
            //if(model.WithDefauls == null)
            //     ((CompuTec.Core2.Beans.IAdvancedUDOBean)udo).SetChangingFromUdo(!(bool)model.WithDefauls);
            var serializer = GetService<ISerializerHandler>().Get(typeof(CompuTec.AppEngine.First.Objects.IRequirement)) as UdoChildBeanSerializer<CompuTec.AppEngine.FirstPlugin.Models.Models.Objects.Requirement, CompuTec.AppEngine.First.Objects.IRequirement>;
            udo.Requirements.Add();
            udo.Requirements.SetCurrentLine(udo.Requirements.Count - 1);
            serializer.FillNew(udo.Requirements, model);
            Update(udo);
            return Ok<CompuTec.AppEngine.FirstPlugin.Models.Models.Objects.Requirement>(serializer.ToModel(udo.Requirements));
        }

        [HttpPut]
        [ODataRoute("({Code})/Requirements({RequirementsLineNum})")]
        public IHttpActionResult PutRequirement([FromODataUri] string Code, [FromODataUri] int RequirementsLineNum, CompuTec.AppEngine.FirstPlugin.Models.Models.Objects.Requirement model)
        {
            var udo = GetObjectInstance(Code);
            var serializer = GetService<ISerializerHandler>().Get(typeof(CompuTec.AppEngine.First.Objects.IRequirement)) as UdoChildBeanSerializer<CompuTec.AppEngine.FirstPlugin.Models.Models.Objects.Requirement, CompuTec.AppEngine.First.Objects.IRequirement>;
            var Requirements = udo.Requirements.FirstOrDefault(item => RequirementsLineNum == item.U_LineNum);
            if (Requirements == null)
                throw new NotFoundException("Requirements not found", "Requirements not found");
            serializer.Update(Requirements, model);
            Update(udo);
            return Ok<CompuTec.AppEngine.FirstPlugin.Models.Models.Objects.Requirement>(serializer.ToModel(Requirements));
        }

        [HttpPatch]
        [ODataRoute("({Code})/Requirements({RequirementsLineNum})")]
        public IHttpActionResult PatchRequirement([FromODataUri] string Code, [FromODataUri] int RequirementsLineNum, DeepDelta<CompuTec.AppEngine.FirstPlugin.Models.Models.Objects.Requirement> model)
        {
            var udo = GetObjectInstance(Code);
            var serializer = GetService<ISerializerHandler>().Get(typeof(CompuTec.AppEngine.First.Objects.IRequirement)) as UdoChildBeanSerializer<CompuTec.AppEngine.FirstPlugin.Models.Models.Objects.Requirement, CompuTec.AppEngine.First.Objects.IRequirement>;
            var Requirements = udo.Requirements.FirstOrDefault(item => RequirementsLineNum == item.U_LineNum);
            if (Requirements == null)
                throw new NotFoundException("Requirements", "Requirements");
            var currentModel = serializer.ToModel(Requirements);
            model.Patch(currentModel);
            serializer.Update(Requirements, currentModel);
            Update(udo);
            return Ok<CompuTec.AppEngine.FirstPlugin.Models.Models.Objects.Requirement>(serializer.ToModel(Requirements));
        }

        [HttpDelete]
        [ODataRoute("({Code})/Requirements({RequirementsLineNum})")]
        public IHttpActionResult DeleteRequirement([FromODataUri] string Code, [FromODataUri] int RequirementsLineNum)
        {
            var udo = GetObjectInstance(Code);
            var udoToDelete = (udo.Requirements as CompuTec.Core2.Beans.IAdvancedUDOChildBean).IMasterBean.Childs.FirstOrDefault(childBean => (childBean as CompuTec.AppEngine.First.Objects.IRequirement).U_LineNum == RequirementsLineNum && (childBean as CompuTec.Core2.Beans.IAdvancedUDOChildBean).IsRowFilled());
            if (udoToDelete != null)
            {
                var position = (udoToDelete as CompuTec.Core2.Beans.IAdvancedUDOChildBean).CurrentPosition;
                udo.Requirements.DelRowAtPos(position);
            }
            else
            {
                throw new NotFoundException("udo", "udo");
            }

            Update(udo);
            return Ok();
        }
    }
}