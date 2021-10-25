using CompuTec.AppEngine.Base.Infrastructure.Controllers.API;
using CompuTec.ProcessForce.API;
using CompuTec.ProcessForce.API.Core;
using CompuTec.ProcessForce.API.Documents.ManufacturingOrder;
using CompuTec.ProcessForce.API.Scheduling;
using CompuTec.ProcessForce.API.Tools;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Http;

namespace CompuTec.AppEngine.FirstPlugin.Controllers.Api
{
    public class PF_SchedulingController : AppEngineSecureController

    {
        [HttpPost]

        [Route("SchduleMor")]
        public IHttpActionResult SchduleMor([FromBody] List<int> AllRelatedMorsDocEntries)
        {
            bool saving = false;
            //iF YOU NEED pf COMPANY PLEASE USE THIS 
            var pfCompany = Session.GetCompany<IProcessForceCompany>();
            var rsc= pfCompany.CreatePFObject(ObjectTypes.Resource);

            // you are already connected
            List<IManufacturingOrder> listOfMorsToBeAdded = BulkUdoConverter.GetBulkObjects<IManufacturingOrder, int>(Session.Token, ObjectTypes.ManufacturingOrder, AllRelatedMorsDocEntries);
            //GetListOfMors
            //You can manipulate Manufacturing orders now by iterating them and injest all the logic
            foreach (var item in listOfMorsToBeAdded)
            {
                item.U_SchedulingMtd = PF_MORSchedulingMthd.Forward;
                item.U_PlannedStartDate = DateTime.Today.AddDays(1);
                item.U_PlannedStartTime = item.U_PlannedStartDate;
            }

            // AllRelatedMorsDocEntries this is a list that contains docentry of MORS to be scheduled on one run
            var sm = new CompuTec.ProcessForce.API.Scheduling.ScheduleManager(Session.Token);

            MultiScheduleParameters param =
                       Activator.CreateInstance(typeof(MultiScheduleParameters),
                       System.Reflection.BindingFlags.NonPublic |
                         System.Reflection.BindingFlags.Instance,
                       null, new object[] { Session.Token }, null) as MultiScheduleParameters;

            listOfMorsToBeAdded.ForEach(m => param.Add(m));
            param.UpdateParents();
            var scheduledMors = sm.Schedule(param);
            //save the mor list scheduledMors
            if (saving)
            {
                foreach (var item in listOfMorsToBeAdded)
                {
                    item.Update();
                }
            }

            return Ok("");
        }
    }
}
