using Enums;
using System.Collections.Generic;
using System.Linq;
using CompuTec.AppEngine.Base.Infrastructure.Controllers.Serialization;
using CompuTec.AppEngine.Base.Infrastructure.Exceptions;
using CompuTec.AppEngine.Base.Infrastructure.Services;

namespace CompuTec.AppEngine.FirstPlugin.Serializer.Serializers.ToDo
{
    public partial class ToDoSerializer : UdoBeanSerializer<CompuTec.AppEngine.FirstPlugin.Models.Models.ToDo.ToDo, CompuTec.AppEngine.FirstPlugin.API.BusinessObjects.ToDo.IToDo>
    {
        public override CompuTec.AppEngine.FirstPlugin.Models.Models.ToDo.ToDo ToModel(CompuTec.AppEngine.FirstPlugin.API.BusinessObjects.ToDo.IToDo udo)
        {
            var model = new CompuTec.AppEngine.FirstPlugin.Models.Models.ToDo.ToDo();
            model.Code = udo.Code;
            model.Name = udo.Name;
            model.U_Deadline = udo.U_Deadline;
            model.U_TaskName = udo.U_TaskName;
            model.U_Description = udo.U_Description;
            model.U_Priority = (ToDoPriority)((int)udo.U_Priority);
            UDFsToModel(udo, model);
            model.Requirements = new List<CompuTec.AppEngine.FirstPlugin.Models.Models.ToDo.ToDoRequirement>();
            udo.Requirements.Where(udoRequirements => (udoRequirements as CompuTec.Core2.Beans.IAdvancedUDOChildBean).IsRowFilled()).ToList().ForEach(udoRequirements =>
            {
                var requirements = new CompuTec.AppEngine.FirstPlugin.Models.Models.ToDo.ToDoRequirement();
                model.Requirements.Add(requirements);
                requirements.U_Name = udoRequirements.U_Name;
                requirements.U_Quantity = udoRequirements.U_Quantity;
                requirements.U_LineNum = udoRequirements.U_LineNum;
                UDFsToModel(udoRequirements, requirements);
            });
            return model;
        }

        public override CompuTec.AppEngine.FirstPlugin.API.BusinessObjects.ToDo.IToDo Update(CompuTec.AppEngine.FirstPlugin.API.BusinessObjects.ToDo.IToDo udo, CompuTec.AppEngine.FirstPlugin.Models.Models.ToDo.ToDo model)
        {
            if (model.Code != null)
                udo.Code = model.Code;
            if (model.Name != null)
                udo.Name = model.Name;
            if (model.U_Deadline != null)
            {
                udo.U_Deadline = (System.DateTime)model.U_Deadline;
            }
            else
            {
                udo.U_Deadline = default(System.DateTime);
            }

            if (model.U_TaskName != null)
                udo.U_TaskName = model.U_TaskName;
            if (model.U_Description != null)
                udo.U_Description = model.U_Description;
            if (model.U_Priority != null)
            {
                udo.U_Priority = (CompuTec.AppEngine.FirstPlugin.API.Enums.ToDoPriority)((int)model.U_Priority);
            }
            else
            {
                udo.U_Priority = default(CompuTec.AppEngine.FirstPlugin.API.Enums.ToDoPriority);
            }

            UDFsToUdo(udo, model);
            if (model.Requirements == null)
            {
                model.Requirements = new List<CompuTec.AppEngine.FirstPlugin.Models.Models.ToDo.ToDoRequirement>();
            }

            var requirementsMasterBean = (udo.Requirements as CompuTec.Core2.Beans.IAdvancedUDOChildBean).IMasterBean.Childs;
            var requirementsToDelete = requirementsMasterBean.Where(childBean => (!model.Requirements.Any(requirements => (childBean as CompuTec.AppEngine.FirstPlugin.API.BusinessObjects.ToDo.IToDoRequirement).U_LineNum == requirements.U_LineNum) || !(childBean as CompuTec.Core2.Beans.IAdvancedUDOChildBean).IsRowFilled())).Select(i => (i as CompuTec.Core2.Beans.IAdvancedUDOChildBean).CurrentPosition).OrderByDescending(i => i);
            foreach (var position in requirementsToDelete)
            {
                udo.Requirements.DelRowAtPos(position);
            }

            model.Requirements.ForEach(requirements =>
            {
                CompuTec.AppEngine.FirstPlugin.API.BusinessObjects.ToDo.IToDoRequirement requirementsItem = null;
                if (requirements.U_LineNum == 0)
                {
                    udo.Requirements.SetCurrentLine(udo.Requirements.Count - 1);
                    if ((udo.Requirements as CompuTec.Core2.Beans.IAdvancedUDOChildBean).IsRowFilled())
                    {
                        udo.Requirements.Add();
                        udo.Requirements.SetCurrentLine(udo.Requirements.Count - 1);
                    }

                    requirementsItem = udo.Requirements;
                }
                else
                {
                    requirementsItem = requirementsMasterBean.FirstOrDefault(childBean => (childBean as CompuTec.AppEngine.FirstPlugin.API.BusinessObjects.ToDo.IToDoRequirement).U_LineNum == requirements.U_LineNum) as CompuTec.AppEngine.FirstPlugin.API.BusinessObjects.ToDo.IToDoRequirement;
                    if (requirementsItem == null)
                        throw new NotFoundException($"CompuTec.AppEngine.FirstPlugin.API.BusinessObjects.ToDo.IToDoRequirement.U_LineNum", $"{requirements.U_LineNum}");
                }

                if (requirements.U_Name != null)
                    requirementsItem.U_Name = requirements.U_Name;
                if (requirements.U_Quantity != null)
                {
                    requirementsItem.U_Quantity = (int)requirements.U_Quantity;
                }
                else
                {
                    requirementsItem.U_Quantity = default(int);
                }

                if (requirements.U_LineNum != null)
                {
                    requirementsItem.U_LineNum = (int)requirements.U_LineNum;
                }
                else
                {
                    requirementsItem.U_LineNum = default(int);
                }

                UDFsToUdo(requirementsItem, requirements);
            });
            return udo;
        }

        public override CompuTec.AppEngine.FirstPlugin.API.BusinessObjects.ToDo.IToDo FillNew(CompuTec.AppEngine.FirstPlugin.API.BusinessObjects.ToDo.IToDo udo, CompuTec.AppEngine.FirstPlugin.Models.Models.ToDo.ToDo model)
        {
            if (model.Code != null)
                udo.Code = model.Code;
            if (model.Name != null)
                udo.Name = model.Name;
            if (model.U_Deadline != null)
                udo.U_Deadline = (System.DateTime)model.U_Deadline;
            if (model.U_TaskName != null)
                udo.U_TaskName = model.U_TaskName;
            if (model.U_Description != null)
                udo.U_Description = model.U_Description;
            if (model.U_Priority != null)
                udo.U_Priority = (CompuTec.AppEngine.FirstPlugin.API.Enums.ToDoPriority)((int)model.U_Priority);
            UDFsToUdo(udo, model);
            if (model.Requirements == null)
            {
                model.Requirements = new List<CompuTec.AppEngine.FirstPlugin.Models.Models.ToDo.ToDoRequirement>();
            }

            var requirementsMasterBean = (udo.Requirements as CompuTec.Core2.Beans.IAdvancedUDOChildBean).IMasterBean.Childs;
            if (!(bool)model.WithDefauls)
            {
                var requirementsToDelete = requirementsMasterBean.Select(i => (i as CompuTec.Core2.Beans.IAdvancedUDOChildBean).CurrentPosition).OrderByDescending(i => i);
                foreach (var position in requirementsToDelete)
                {
                    udo.Requirements.DelRowAtPos(position);
                }
            }

            model.Requirements.ForEach(requirements =>
            {
                udo.Requirements.SetCurrentLine(udo.Requirements.Count - 1);
                if ((udo.Requirements as CompuTec.Core2.Beans.IAdvancedUDOChildBean).IsRowFilled())
                {
                    udo.Requirements.Add();
                    udo.Requirements.SetCurrentLine(udo.Requirements.Count - 1);
                }

                if (requirements.U_Name != null)
                    udo.Requirements.U_Name = requirements.U_Name;
                if (requirements.U_Quantity != null)
                    udo.Requirements.U_Quantity = (int)requirements.U_Quantity;
                if (requirements.U_LineNum != null)
                    udo.Requirements.U_LineNum = (int)requirements.U_LineNum;
                UDFsToUdo(udo.Requirements, requirements);
            });
            return udo;
        }

        protected override CompuTec.AppEngine.FirstPlugin.API.BusinessObjects.ToDo.IToDo FillNewExtended(CompuTec.AppEngine.FirstPlugin.API.BusinessObjects.ToDo.IToDo udo, CompuTec.AppEngine.FirstPlugin.Models.Models.ToDo.ToDo model)
        {
            if (model.Code != null)
                udo.Code = model.Code;
            if (model.Name != null)
                udo.Name = model.Name;
            if (model.U_Deadline != null)
                udo.U_Deadline = (System.DateTime)model.U_Deadline;
            if (model.U_TaskName != null)
                udo.U_TaskName = model.U_TaskName;
            if (model.U_Description != null)
                udo.U_Description = model.U_Description;
            if (model.U_Priority != null)
                udo.U_Priority = (CompuTec.AppEngine.FirstPlugin.API.Enums.ToDoPriority)((int)model.U_Priority);
            UDFsToUdo(udo, model);
            if (model.Requirements == null)
            {
                model.Requirements = new List<CompuTec.AppEngine.FirstPlugin.Models.Models.ToDo.ToDoRequirement>();
            }

            var requirementsMasterBean = (udo.Requirements as CompuTec.Core2.Beans.IAdvancedUDOChildBean).IMasterBean.Childs;
            if (!(bool)model.WithDefauls)
            {
                var requirementsToDelete = requirementsMasterBean.Select(i => (i as CompuTec.Core2.Beans.IAdvancedUDOChildBean).CurrentPosition).OrderByDescending(i => i);
                foreach (var position in requirementsToDelete)
                {
                    udo.Requirements.DelRowAtPos(position);
                }
            }

            model.Requirements.ForEach(requirements =>
            {
                udo.Requirements.SetCurrentLine(udo.Requirements.Count - 1);
                if ((udo.Requirements as CompuTec.Core2.Beans.IAdvancedUDOChildBean).IsRowFilled())
                {
                    udo.Requirements.Add();
                    udo.Requirements.SetCurrentLine(udo.Requirements.Count - 1);
                }

                if (requirements.U_Name != null)
                    udo.Requirements.U_Name = requirements.U_Name;
                if (requirements.U_Quantity != null)
                    udo.Requirements.U_Quantity = (int)requirements.U_Quantity;
                if (requirements.U_LineNum != null)
                    udo.Requirements.U_LineNum = (int)requirements.U_LineNum;
                UDFsToUdo(udo.Requirements, requirements);
            });
            return udo;
        }
    }
}