using CompuTec.AppEngine.JobsSamplePlugin.Models.Todo;
using System;
using System.Collections.Generic;

namespace CompuTec.AppEngine.JobsSamplePlugin.Api.Services
{
    public class DataSource
    {
        private List<Task> _tasks = new List<Task>()
        {
            new Task(){ UserName = "manager", CreationDate = DateTime.Now.AddHours(-2), Title = "Get my mac fixed", Id = Guid.NewGuid().ToString()},
            new Task(){ UserName = "manager", CreationDate = DateTime.Now.AddHours(-1), Title = "Book tickets to Hogwart", Id = Guid.NewGuid().ToString()},
            new Task(){ UserName = "manager", CreationDate = DateTime.Now.AddMinutes(-30), Title = "Talk with Michael about new tasks", Id = Guid.NewGuid().ToString()},
        };

        public List<Task> TodoDataSource => _tasks;
    }
}
