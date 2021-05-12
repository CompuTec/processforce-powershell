using System;
using System.Collections.Generic;
using System.Web.Http;
using System.Web.Http.Description;
using CompuTec.AppEngine.Base.Infrastructure.Annotation;
using CompuTec.AppEngine.Base.Infrastructure.Controllers;
using CompuTec.AppEngine.Base.Infrastructure.Controllers.API;
using CompuTec.AppEngine.Base.Infrastructure.Controllers.OData.Delta;

namespace CompuTec.AppEngine.FirstPlugin.Controllers.Api
{
    public partial class ToDoController : AppEngineCore2Controller<CompuTec.AppEngine.FirstPlugin.Models.Models.Objects.ToDo, CompuTec.AppEngine.First.Objects.IToDo, string>
    {
        protected override string TableName => "@SAMPLE_OTDO";
        protected override string KeyName => "Code";
        protected override string ObjectType => "Sample_ToDo";
        protected override eUDOVersion UDOVersion => eUDOVersion.UDO_20;
    }
}