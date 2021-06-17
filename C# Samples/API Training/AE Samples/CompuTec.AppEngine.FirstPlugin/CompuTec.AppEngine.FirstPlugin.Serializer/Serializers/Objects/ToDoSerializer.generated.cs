using System.Collections.Generic;
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
            model.U_Deadline = udo.U_Deadline;
            model.U_TaskName = udo.U_TaskName;
            model.U_Description = udo.U_Description;
            model.U_Priority = udo.U_Priority;
            UDFsToModel(udo, model);
            model.Requirements = new List<CompuTec.AppEngine.FirstPlugin.Models.Models.Objects.Requirement>();
            udo.Requirements.Where(udoRequirements => (udoRequirements as CompuTec.Core2.Beans.IAdvancedUDOChildBean).IsRowFilled()).ToList().ForEach(udoRequirements =>
            {
                var requirements = new CompuTec.AppEngine.FirstPlugin.Models.Models.Objects.Requirement();
                model.Requirements.Add(requirements);
                requirements.Code = udoRequirements.Code;
                requirements.Name = udoRequirements.Name;
                requirements.U_LineNum = udoRequirements.U_LineNum;
                UDFsToModel(udoRequirements, requirements);
            });
            return model;
        }

        public override CompuTec.AppEngine.First.Objects.IToDo Update(CompuTec.AppEngine.First.Objects.IToDo udo, CompuTec.AppEngine.FirstPlugin.Models.Models.Objects.ToDo model)
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
                udo.U_Priority = model.U_Priority;
            UDFsToUdo(udo, model);
            if (model.Requirements == null)
            {
                model.Requirements = new List<CompuTec.AppEngine.FirstPlugin.Models.Models.Objects.Requirement>();
            }

            var requirementsMasterBean = (udo.Requirements as CompuTec.Core2.Beans.IAdvancedUDOChildBean).IMasterBean.Childs;
            var requirementsToDelete = requirementsMasterBean.Where(childBean => (!model.Requirements.Any(requirements => (childBean as CompuTec.AppEngine.First.Objects.IRequirement).U_LineNum == requirements.U_LineNum) || !(childBean as CompuTec.Core2.Beans.IAdvancedUDOChildBean).IsRowFilled())).Select(i => (i as CompuTec.Core2.Beans.IAdvancedUDOChildBean).CurrentPosition).OrderByDescending(i => i);
            foreach (var position in requirementsToDelete)
            {
                udo.Requirements.DelRowAtPos(position);
            }

            model.Requirements.ForEach(requirements =>
            {
                CompuTec.AppEngine.First.Objects.IRequirement requirementsItem = null;
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
                    requirementsItem = requirementsMasterBean.FirstOrDefault(childBean => (childBean as CompuTec.AppEngine.First.Objects.IRequirement).U_LineNum == requirements.U_LineNum) as CompuTec.AppEngine.First.Objects.IRequirement;
                    if (requirementsItem == null)
                        throw new NotFoundException($"CompuTec.AppEngine.First.Objects.IRequirement.U_LineNum", $"{requirements.U_LineNum}");
                }

                if (requirements.Code != null)
                    requirementsItem.Code = requirements.Code;
                if (requirements.Name != null)
                    requirementsItem.Name = requirements.Name;
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

        public override CompuTec.AppEngine.First.Objects.IToDo FillNew(CompuTec.AppEngine.First.Objects.IToDo udo, CompuTec.AppEngine.FirstPlugin.Models.Models.Objects.ToDo model)
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
                udo.U_Priority = model.U_Priority;
            UDFsToUdo(udo, model);
            if (model.Requirements == null)
            {
                model.Requirements = new List<CompuTec.AppEngine.FirstPlugin.Models.Models.Objects.Requirement>();
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

                if (requirements.Code != null)
                    udo.Requirements.Code = requirements.Code;
                if (requirements.Name != null)
                    udo.Requirements.Name = requirements.Name;
                if (requirements.U_LineNum != null)
                    udo.Requirements.U_LineNum = (int)requirements.U_LineNum;
                UDFsToUdo(udo.Requirements, requirements);
            });
            return udo;
        }

        protected override CompuTec.AppEngine.First.Objects.IToDo FillNewExtended(CompuTec.AppEngine.First.Objects.IToDo udo, CompuTec.AppEngine.FirstPlugin.Models.Models.Objects.ToDo model)
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
                udo.U_Priority = model.U_Priority;
            UDFsToUdo(udo, model);
            if (model.Requirements == null)
            {
                model.Requirements = new List<CompuTec.AppEngine.FirstPlugin.Models.Models.Objects.Requirement>();
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

                if (requirements.Code != null)
                    udo.Requirements.Code = requirements.Code;
                if (requirements.Name != null)
                    udo.Requirements.Name = requirements.Name;
                if (requirements.U_LineNum != null)
                    udo.Requirements.U_LineNum = (int)requirements.U_LineNum;
                UDFsToUdo(udo.Requirements, requirements);
            });
            return udo;
        }
    }
}