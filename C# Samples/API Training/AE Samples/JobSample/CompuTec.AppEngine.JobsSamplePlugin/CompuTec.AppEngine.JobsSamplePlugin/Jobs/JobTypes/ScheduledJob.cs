using CompuTec.AppEngine.Base.Infrastructure.Jobs;
using CompuTec.AppEngine.Base.Infrastructure.Jobs.Annotations;
using CompuTec.AppEngine.Base.Infrastructure.Security;
using CompuTec.Core2;
using CompuTec.Core2.Connection;
using StructureMap;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CompuTec.AppEngine.JobsSamplePlugin.Jobs.JobTypes
{
    [RecurringJob(JobId = "ScheduledTestJob", CronExpression = "*/30 * * * * *", Description = "Sample Schedue based job Job")]
    public class ScheduledJob : SecureJob
    {
        public ScheduledJob(Session session, IContainer container) : base(session, container)
        {
        }
        public override void Call()
        {
            //get logger 
            var log = Container.GetInstance<NLog.Logger>();
            try
            {
                //please notice that this job is posted with info message please go to traceConfiguration in adminstration panel and change log level to info
                log.Info($"{nameof(ScheduledJob)} job started");
                var itemcode = $"SJ_{DateTime.Now.ToString("yyyyyMMddHHmmss")}";
                var company = ConnectionHolder.GetCompany(Session.Token);
                var itm = company.GetBusinessObject(BaseLayer.DI.BoObjectTypes.oItems) as BaseLayer.DI.Items;
                itm.ItemCode = itemcode;
                itm.ItemName = $"EppEngine Item {itemcode}";
                var added= itm.Add();
                if(added==0)
                {
                    log.Info($"{nameof(ScheduledJob)} job created an item {itemcode}");
                }else
                {
                    log.Error($"{nameof(ScheduledJob)} failed to add an item {itemcode} : {company.GetLastErrorDescription()}");
                }
            }
            catch(Exception ex)
            {
                log.Error(ex, $"{nameof(ScheduledJob)} throws exception");
            }

        }
    }
}
