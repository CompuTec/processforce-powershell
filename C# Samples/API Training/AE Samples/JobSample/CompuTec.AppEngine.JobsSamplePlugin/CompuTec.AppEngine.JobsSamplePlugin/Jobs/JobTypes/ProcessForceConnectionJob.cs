using CompuTec.AppEngine.Base.Infrastructure.Jobs;
using CompuTec.AppEngine.Base.Infrastructure.Jobs.Annotations;
using CompuTec.Core.DI.Database;
using CompuTec.ProcessForce.API;
using Newtonsoft.Json;
using NLog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CompuTec.AppEngine.JobsSamplePlugin.Jobs.JobTypes
{
    [EventBusJob(JobId = "CreateMORFromSalesSample", Description = "Create MO on SO Creation", ContentType = "17", ActionType = "A")]
    public class ProcessForceConnectionJob : EventBusSecureJob
    {
        private readonly Logger log;

        public  ProcessForceConnectionJob(Base.Infrastructure.Security.Session session, StructureMap.IContainer container, EventBus.Message message) : base(session, container, message)
        {
            log = Container.GetInstance<NLog.Logger>();
        }

        public override void Call()
        {
            
            try
            {
                dynamic body = JsonConvert.DeserializeObject(Message.Body);
                int docEntry = body.DocEntry;
                log.Info($"{nameof(ProcessForceConnectionJob)} job started for Document :{docEntry}");
                var company = Session.GetCompany<IProcessForceCompany>();
                var qm = new QueryManager();
                qm.CommandText = @"Select t0.""Code"", t1.""Quantity"" from ""@CT_PF_OBOM"" t0 
inner join RDR1 t1 on t0.""U_ItemCode"" =t1.""ItemCode"" and t0.""U_Revision""=t1.""U_Revision"" 
where t1.""DocEntry"" = @DocEntry";
                qm.AddParameter("DocEntry", docEntry);
                using (var rs=qm.Execute(Session.Token))
                {
                    for (int i = 0; i < rs.RecordCount; i++)
                    {
                        string code = rs.Fields.Item("Code").Value;
                        double qty = rs.Fields.Item("Quantity").Value;
                        CreateMor(code, qty);
                        rs.MoveNext();
                    }
                }
            }
            catch (Exception ex)
            {
                log.Error(ex, $"{nameof(ProcessForceConnectionJob)} failed {ex.Message}");
            }
        }

        private void CreateMor(string bomCode, dynamic quantity)
        {
            var mor = Session.GetCompany<IProcessForceCompany>().CreatePFObject(ProcessForce.API.Core.ObjectTypes.ManufacturingOrder) as ProcessForce.API.Documents.ManufacturingOrder.IManufacturingOrder;
            mor.U_BOMCode = bomCode;
            mor.U_Quantity = quantity;
            mor.U_Status = ProcessForce.API.Enumerators.ManufacturingOrderStatus.NotScheduled;
            var added = mor.Add();
            if(added==0)
            {
                log.Info($"MOR added {bomCode},{quantity} ");
            }else
            {
                log.Error($"MOR added not added {mor.ErrorDescription}");
            }

        }
    }
}
