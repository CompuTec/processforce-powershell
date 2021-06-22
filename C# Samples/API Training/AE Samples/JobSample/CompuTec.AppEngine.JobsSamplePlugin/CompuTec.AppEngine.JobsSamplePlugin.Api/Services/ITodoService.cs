using CompuTec.AppEngine.JobsSamplePlugin.Models.Todo;
using System.Collections.Generic;

namespace CompuTec.AppEngine.JobsSamplePlugin.Api.Services
{
    public interface ITodoService
    {
        IEnumerable<Task> GetAll();


        Task Get(string id);

        Task Add(Task task);

        Task Update(string id, Task task);


        void Delete(string id);
    }
}