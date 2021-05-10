﻿using System.Collections.Generic;
using System.Linq;
using CompuTec.AppEngine.Base.Infrastructure.Controllers.Serialization;
using CompuTec.AppEngine.Base.Infrastructure.Exceptions;
using CompuTec.AppEngine.Base.Infrastructure.Services;

namespace CompuTec.AppEngine.FirstPlugin.Serializer.Serializers.Objects
{
    public partial class ToDoSerializer : UdoBeanSerializer<CompuTec.AppEngine.FirstPlugin.Models.Models.Objects.ToDo, CompuTec.AppEngine.First.Objects.IToDo>
    {
        public override CompuTec.AppEngine.FirstPlugin.Models.Models.Objects.ToDo ToModel(CompuTec.AppEngine.First.Objects.IToDo udo)
        {
            var model = new CompuTec.AppEngine.FirstPlugin.Models.Models.Objects.ToDo();
            model.Code = udo.Code;
            model.Name = udo.Name;
            UDFsToModel(udo, model);
            return model;
        }

        public override CompuTec.AppEngine.First.Objects.IToDo Update(CompuTec.AppEngine.First.Objects.IToDo udo, CompuTec.AppEngine.FirstPlugin.Models.Models.Objects.ToDo model)
        {
            udo.Code = model.Code;
            udo.Name = model.Name;
            UDFsToUdo(udo, model);
            return udo;
        }

        public override CompuTec.AppEngine.First.Objects.IToDo FillNew(CompuTec.AppEngine.First.Objects.IToDo udo, CompuTec.AppEngine.FirstPlugin.Models.Models.Objects.ToDo model)
        {
            if (model.Code != null)
                udo.Code = model.Code;
            if (model.Name != null)
                udo.Name = model.Name;
            UDFsToUdo(udo, model);
            return udo;
        }

        protected override CompuTec.AppEngine.First.Objects.IToDo FillNewExtended(CompuTec.AppEngine.First.Objects.IToDo udo, CompuTec.AppEngine.FirstPlugin.Models.Models.Objects.ToDo model)
        {
            if (model.Code != null)
                udo.Code = model.Code;
            if (model.Name != null)
                udo.Name = model.Name;
            UDFsToUdo(udo, model);
            return udo;
        }
    }
}