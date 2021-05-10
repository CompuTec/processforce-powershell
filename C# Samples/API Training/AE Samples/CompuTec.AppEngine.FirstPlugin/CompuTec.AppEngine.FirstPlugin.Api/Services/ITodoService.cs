using CompuTec.AppEngine.FirstPlugin.Models.Todo;
using System.Collections.Generic;

namespace CompuTec.AppEngine.FirstPlugin.Api.Services
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