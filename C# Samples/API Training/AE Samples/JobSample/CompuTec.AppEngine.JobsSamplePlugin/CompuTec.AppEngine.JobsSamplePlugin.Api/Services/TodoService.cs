using CompuTec.AppEngine.Base.Infrastructure.Security;
using CompuTec.AppEngine.Base.Infrastructure.Services;
using CompuTec.AppEngine.JobsSamplePlugin.Models.Todo;
using StructureMap;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CompuTec.AppEngine.JobsSamplePlugin.Api.Services
{
    public class TodoService : BaseSecureService, ITodoService
    {
        public TodoService(Session session, IContainer container) : base(session, container)
        {
        }


        public IEnumerable<Task> GetAll()
        {
            return Container.GetInstance<DataSource>().TodoDataSource.Where(t => t.UserName.Equals(Session.UserName));
        }

        public Task Get(string id)
        {
            return Container.GetInstance<DataSource>().TodoDataSource.FirstOrDefault(t => t.Id.Equals(id) && t.UserName.Equals(Session.UserName));
        }

        public Task Add(Task task)
        {
            var newTask = new Task()
            {
                Id = Guid.NewGuid().ToString(),
                UserName = Session.UserName,
                Title = task.Title,
                CreationDate = DateTime.Now,
                IsDone = task.IsDone
            };

            Container.GetInstance<DataSource>().TodoDataSource.Add(newTask);

            return newTask;
        }

        public Task Update(string id, Task task)
        {
            var currentTask = Container.GetInstance<DataSource>().TodoDataSource.First(t => t.Id.Equals(id));
            currentTask.Title = task.Title;
            currentTask.IsDone = task.IsDone;

            return currentTask;
        }

        public void Delete(string id)
        {
            Container.GetInstance<DataSource>().TodoDataSource.RemoveAll(t => t.Id.Equals(id));
        }
    }
}
