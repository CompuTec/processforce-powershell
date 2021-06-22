using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CompuTec.AppEngine.Base.Infrastructure.Jobs;
using CompuTec.AppEngine.Base.Infrastructure.Jobs.Annotations;
using CompuTec.Core2.Connection;

namespace CompuTec.AppEngine.JobsSamplePlugin.Jobs.JobTypes
{
    [BackgroundJob("SampleOneTimeJob", "One Time Sample Job")]
    public class OneTimeJob : CompuTec.AppEngine.Base.Infrastructure.Jobs.SecureJob

    {
        public OneTimeJob(Base.Infrastructure.Security.Session session, StructureMap.IContainer container) : base(session, container)
        {
        }

        public override void Call()
        {
            var log = Container.GetInstance<NLog.Logger>(); try
            {

                log.Info($"{nameof(OneTimeJob)} job started for item:");
                var company = ConnectionHolder.GetCompany(Session.Token);
                var bp = company.GetBusinessObject(BaseLayer.DI.BoObjectTypes.oBusinessPartners) as BaseLayer.DI.BusinessPartners;
                var cardCode= $"OTJ_{DateTime.Now.ToString("yyyyyMMddHHmmss")}";
                bp.CardCode = cardCode;
                bp.CardType = BaseLayer.DI.BoCardTypes.cCustomer;
                var added = bp.Add();
                if (added == 0)
                {
                    log.Info($"{nameof(OneTimeJob)} job created an BP {cardCode}");
                }
                else
                {
                    log.Error($"{nameof(OneTimeJob)} failed to add an BP {cardCode} : {company.GetLastErrorDescription()}");
                }
            }
            catch (Exception ex)
            {

                log.Error(ex, $"{nameof(OneTimeJob)} throws exception {ex.Message}");

            }
        }
    }
}
