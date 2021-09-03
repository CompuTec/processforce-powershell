using Enums;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using CompuTec.AppEngine.Base.Infrastructure.Controllers;
using System.ComponentModel.DataAnnotations.Schema;
using Microsoft.AspNet.OData.Builder;

namespace CompuTec.AppEngine.FirstPlugin.Models.Models.ToDo
{
    public class ToDo : AppEngineUdo, ICloneable
    {
        [Key]
        public string Code { get; set; }

        public string Name { get; set; }

        public System.DateTime? U_Deadline { get; set; }

        public string U_TaskName { get; set; }

        public string U_Description { get; set; }

        public ToDoPriority? U_Priority { get; set; }

        [Contained]
        public List<ToDoRequirement> Requirements { get; set; }

        object ICloneable.Clone()
        {
            return (ToDo)this.MemberwiseClone();
        }

        public ToDo Clone()
        {
            return (ToDo)this.MemberwiseClone();
        }
    }
}