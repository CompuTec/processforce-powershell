using Enums;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using CompuTec.AppEngine.Base.Infrastructure.Controllers;
using System.ComponentModel.DataAnnotations.Schema;

namespace CompuTec.AppEngine.FirstPlugin.Models.Models.ToDo
{
    public class ToDoRequirement : AppEngineUdo, ICloneable
    {
        public string U_Name { get; set; }

        public int? U_Quantity { get; set; }

        [Key]
        public int? U_LineNum { get; set; }

        object ICloneable.Clone()
        {
            return (ToDoRequirement)this.MemberwiseClone();
        }

        public ToDoRequirement Clone()
        {
            return (ToDoRequirement)this.MemberwiseClone();
        }
    }
}