using CompuTec.AppEngine.Base.Infrastructure.Jobs;
using CompuTec.AppEngine.Base.Infrastructure.Jobs.Annotations;
using CompuTec.AppEngine.Base.Infrastructure.Security;
using CompuTec.AppEngine.EventBus;
using CompuTec.Core2.Connection;
using Newtonsoft.Json;
using StructureMap;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CompuTec.AppEngine.JobsSamplePlugin.Jobs.JobTypes
{
	[EventBusJob(JobId = "EventBasedSampleJob", Description = "Sample Event Based Job", ContentType = "4", ActionType = "A")]
    public class EventBasedJob : EventBusSecureJob
    {
		public EventBasedJob(Session session, IContainer container, Message message) : base(session, container, message)
		{
		}

        public override void Call()
        {
            var log = Container.GetInstance<NLog.Logger>();
            try
            {
                dynamic body = JsonConvert.DeserializeObject(Message.Body);
                string itemCode = body.ItemCode;
                log.Info($"{nameof(EventBasedJob)} job started for item:{itemCode}");
                var company = ConnectionHolder.GetCompany(Session.Token);
                var item = company.GetBusinessObject(BaseLayer.DI.BoObjectTypes.oItems) as BaseLayer.DI.Items;
                item.GetByKey(itemCode);
                item.User_Text = "Add action was catched by AE job and updated ";
                item.Update();
            }
            catch(Exception ex)
            {
                log.Error(ex, $"{nameof(EventBasedJob)} failed {ex.Message}");
            }

        }
    }
}
