using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using CompuTec.AppEngine.Base.Infrastructure.Controllers;
using System.ComponentModel.DataAnnotations.Schema;
using System.Web.OData.Builder;

namespace CompuTec.AppEngine.FirstPlugin.Models.Models.Objects
{
    public class ToDo : AppEngineUdo, ICloneable
    {
        [Key]
        public string Code { get; set; }

        public string Name { get; set; }

        public System.DateTime? U_Deadline { get; set; }

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