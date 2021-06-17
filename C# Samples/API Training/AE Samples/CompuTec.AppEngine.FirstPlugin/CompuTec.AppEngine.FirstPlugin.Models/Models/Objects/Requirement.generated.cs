using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using CompuTec.AppEngine.Base.Infrastructure.Controllers;
using System.ComponentModel.DataAnnotations.Schema;
using System.Web.OData.Builder;

namespace CompuTec.AppEngine.FirstPlugin.Models.Models.Objects
{
    public class Requirement : AppEngineUdo, ICloneable
    {
        public string Code { get; set; }

        public string Name { get; set; }

        [Key]
        public int? U_LineNum { get; set; }

        object ICloneable.Clone()
        {
            return (Requirement)this.MemberwiseClone();
        }

        public Requirement Clone()
        {
            return (Requirement)this.MemberwiseClone();
        }
    }
}