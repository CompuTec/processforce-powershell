using CompuTec.AppEngine.Base.Infrastructure.Controllers;
using System;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace CompuTec.AppEngine.FirstPlugin.Models.Todo
{
    public class Task : AppEngineEntity
    {
        [Key]
        public string Id { get; set; }

        [ReadOnly(true)]
        public DateTime CreationDate { get; set; }

        [ReadOnly(true)]
        public string UserName { get; set; }

        public string Title { get; set; }

        public bool IsDone { get; set; }
    }
}
