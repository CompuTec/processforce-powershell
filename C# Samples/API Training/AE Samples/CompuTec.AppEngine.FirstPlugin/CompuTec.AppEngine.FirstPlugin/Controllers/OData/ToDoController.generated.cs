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
    public partial class ToDoController : AppEngineCore2ODataBatchController<CompuTec.AppEngine.FirstPlugin.Models.Models.ToDo.ToDo, CompuTec.AppEngine.FirstPlugin.API.BusinessObjects.ToDo.IToDo, string>
    {
        protected override string TableName => "@SAMPLE_OTDO";
        protected override string KeyName => "Code";
        protected override string ObjectType => "SAMPLE_TO_DO";
        protected override eUDOVersion UDOVersion => eUDOVersion.UDO_20;
        [HttpGet]
        [EnableQuery(MaxExpansionDepth = 10)]
        [ODataRoute("({Code})/Requirements")]
        public IQueryable<CompuTec.AppEngine.FirstPlugin.Models.Models.ToDo.ToDoRequirement> GetToDoRequirement([FromODataUri] string Code)
        {
            var udo = GetObjectInstance(Code);
            var model = Serializer.ToModel(udo);
            var Requirements = model.Requirements;
            return Requirements.AsQueryable<CompuTec.AppEngine.FirstPlugin.Models.Models.ToDo.ToDoRequirement>();
        }

        [HttpGet]
        [EnableQuery(MaxExpansionDepth = 10)]
        [ODataRoute("({Code})/Requirements({RequirementsLineNum})")]
        public SingleResult<CompuTec.AppEngine.FirstPlugin.Models.Models.ToDo.ToDoRequirement> GetToDoRequirement([FromODataUri] string Code, [FromODataUri] int RequirementsLineNum)
        {
            var udo = GetObjectInstance(Code);
            var model = Serializer.ToModel(udo);
            var Requirements = model.Requirements.FirstOrDefault(item => RequirementsLineNum == item.U_LineNum);
            if (Requirements == null)
                throw new NotFoundException("Requirements", "Requirements");
            return SingleResult.Create(new List<CompuTec.AppEngine.FirstPlugin.Models.Models.ToDo.ToDoRequirement>()
            {Requirements}.AsQueryable());
        }

        [HttpPost]
        [ODataRoute("({Code})/Requirements")]
        public IHttpActionResult PostToDoRequirement([FromODataUri] string Code, CompuTec.AppEngine.FirstPlugin.Models.Models.ToDo.ToDoRequirement model)
        {
            var udo = GetObjectInstance(Code);
            //if(model.WithDefauls == null)
            //     ((CompuTec.Core2.Beans.IAdvancedUDOBean)udo).SetChangingFromUdo(!(bool)model.WithDefauls);
            var serializer = GetService<ISerializerHandler>().Get(typeof(CompuTec.AppEngine.FirstPlugin.API.BusinessObjects.ToDo.IToDoRequirement)) as UdoChildBeanSerializer<CompuTec.AppEngine.FirstPlugin.Models.Models.ToDo.ToDoRequirement, CompuTec.AppEngine.FirstPlugin.API.BusinessObjects.ToDo.IToDoRequirement>;
            udo.Requirements.Add();
            udo.Requirements.SetCurrentLine(udo.Requirements.Count - 1);
            serializer.FillNew(udo.Requirements, model);
            Update(udo);
            return Ok<CompuTec.AppEngine.FirstPlugin.Models.Models.ToDo.ToDoRequirement>(serializer.ToModel(udo.Requirements));
        }

        [HttpPut]
        [ODataRoute("({Code})/Requirements({RequirementsLineNum})")]
        public IHttpActionResult PutToDoRequirement([FromODataUri] string Code, [FromODataUri] int RequirementsLineNum, CompuTec.AppEngine.FirstPlugin.Models.Models.ToDo.ToDoRequirement model)
        {
            var udo = GetObjectInstance(Code);
            var serializer = GetService<ISerializerHandler>().Get(typeof(CompuTec.AppEngine.FirstPlugin.API.BusinessObjects.ToDo.IToDoRequirement)) as UdoChildBeanSerializer<CompuTec.AppEngine.FirstPlugin.Models.Models.ToDo.ToDoRequirement, CompuTec.AppEngine.FirstPlugin.API.BusinessObjects.ToDo.IToDoRequirement>;
            var Requirements = udo.Requirements.FirstOrDefault(item => RequirementsLineNum == item.U_LineNum);
            if (Requirements == null)
                throw new NotFoundException("Requirements not found", "Requirements not found");
            serializer.Update(Requirements, model);
            Update(udo);
            return Ok<CompuTec.AppEngine.FirstPlugin.Models.Models.ToDo.ToDoRequirement>(serializer.ToModel(Requirements));
        }

        [HttpPatch]
        [ODataRoute("({Code})/Requirements({RequirementsLineNum})")]
        public IHttpActionResult PatchToDoRequirement([FromODataUri] string Code, [FromODataUri] int RequirementsLineNum, DeepDelta<CompuTec.AppEngine.FirstPlugin.Models.Models.ToDo.ToDoRequirement> model)
        {
            var udo = GetObjectInstance(Code);
            var serializer = GetService<ISerializerHandler>().Get(typeof(CompuTec.AppEngine.FirstPlugin.API.BusinessObjects.ToDo.IToDoRequirement)) as UdoChildBeanSerializer<CompuTec.AppEngine.FirstPlugin.Models.Models.ToDo.ToDoRequirement, CompuTec.AppEngine.FirstPlugin.API.BusinessObjects.ToDo.IToDoRequirement>;
            var Requirements = udo.Requirements.FirstOrDefault(item => RequirementsLineNum == item.U_LineNum);
            if (Requirements == null)
                throw new NotFoundException("Requirements", "Requirements");
            var currentModel = serializer.ToModel(Requirements);
            model.Patch(currentModel);
            serializer.Update(Requirements, currentModel);
            Update(udo);
            return Ok<CompuTec.AppEngine.FirstPlugin.Models.Models.ToDo.ToDoRequirement>(serializer.ToModel(Requirements));
        }

        [HttpDelete]
        [ODataRoute("({Code})/Requirements({RequirementsLineNum})")]
        public IHttpActionResult DeleteToDoRequirement([FromODataUri] string Code, [FromODataUri] int RequirementsLineNum)
        {
            var udo = GetObjectInstance(Code);
            var udoToDelete = (udo.Requirements as CompuTec.Core2.Beans.IAdvancedUDOChildBean).IMasterBean.Childs.FirstOrDefault(childBean => (childBean as CompuTec.AppEngine.FirstPlugin.API.BusinessObjects.ToDo.IToDoRequirement).U_LineNum == RequirementsLineNum && (childBean as CompuTec.Core2.Beans.IAdvancedUDOChildBean).IsRowFilled());
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