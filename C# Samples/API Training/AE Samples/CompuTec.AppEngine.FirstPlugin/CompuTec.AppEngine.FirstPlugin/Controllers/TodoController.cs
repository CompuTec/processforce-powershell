using CompuTec.AppEngine.Base.Infrastructure.Controllers.OData;
using CompuTec.AppEngine.Base.Infrastructure.Controllers.OData.Delta;
using CompuTec.AppEngine.FirstPlugin.Api.Services;
using CompuTec.AppEngine.FirstPlugin.Models.Todo;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Http;
using System.Web.OData;

namespace CompuTec.AppEngine.FirstPlugin.Controllers
{
    public class TodoController : AppEngineODataSecureController<Task>
    {
        [EnableQuery()]
        [HttpGet]
        public IEnumerable<Task> Get()
        {
            var todoService = GetService<ITodoService>();
            return todoService.GetAll();
        }

        [HttpGet]
        [EnableQuery()]
        public SingleResult<Task> Get(string key)
        {
            var todoService = GetService<ITodoService>();

            return SingleResult.Create(new List<Task>() { todoService.Get(key) }.AsQueryable());
        }


        [HttpPost]
        public IHttpActionResult Post(Task add)
        {
            var todoService = GetService<ITodoService>();

            return Ok<Task>(todoService.Add(add));
        }


        [HttpPut]
        public IHttpActionResult Put(string key, Task task)
        {
            var todoService = GetService<ITodoService>();

            return Ok<Task>(todoService.Update(key, task));
        }


        [HttpPatch]
        public IHttpActionResult Patch(string key, DeepDelta<Task> patch)
        {
            var todoService = GetService<ITodoService>();

            var oldTask = todoService.Get(key);
            patch.Patch(oldTask);

            return Ok<Task>(todoService.Update(key, oldTask));
        }


        [HttpDelete]
        public IHttpActionResult Delete(string key)
        {
            var todoService = GetService<ITodoService>();
            todoService.Delete(key);
            return Ok();
        }
    }
}