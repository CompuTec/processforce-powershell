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
        protected override string TableName => "@Sample_OTDO";
        protected override string KeyName => "Code";
        protected override string ObjectType => "Sample_ToDo";
        protected override eUDOVersion UDOVersion => eUDOVersion.UDO_20;
    }
}