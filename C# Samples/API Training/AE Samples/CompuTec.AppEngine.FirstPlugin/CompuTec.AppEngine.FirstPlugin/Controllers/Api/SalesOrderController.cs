using CompuTec.AppEngine.Base.Infrastructure.Controllers.API;
using CompuTec.AppEngine.Base.Infrastructure.Security;
using CompuTec.AppEngine.FirstPlugin.Models.Models.SalesOrder;
using CompuTec.BaseLayer.DI;
using CompuTec.Core2.DI.Database;
using System;
using System.Web.Http;

namespace CompuTec.AppEngine.FirstPlugin.Controllers.Api
{
	public class SalesOrderController : AppEngineSecureController
	{

		
		[HttpPost]
		[Route("AddAttachment")]
		public IHttpActionResult AddAttachment([FromBody] SalesOrderAttachment attachment)
		{
			Documents so = GetSalesOrder(attachment.DocEntry);
			Attachments2 atc = GetAttachment(so.AttachmentEntry);

			bool firstAttachment = false;

			if (atc.AbsoluteEntry == 0)
				firstAttachment = true;

			string AtcPath = GetAttachmentsPath();
		
			atc.Lines.SetCurrentLine(atc.Lines.Count - 1);
			if (!string.IsNullOrWhiteSpace(atc.Lines.FileName))
				atc.Lines.Add();
			atc.Lines.SourcePath = AtcPath.TrimEnd('\\');
			atc.Lines.FileName = attachment.FileName;
			atc.Lines.FileExtension = attachment.FileExtension;

			int res;
			if (firstAttachment)
				res = atc.Add();
			else
				res = atc.Update();

			if (res != 0)
				throw new Exception($"Adding attachment failed: {Company.GetLastErrorDescription()}");

			if (firstAttachment)
			{
				so.AttachmentEntry = Int32.Parse(Company.GetNewObjectKey());
				if (so.Update() != 0)
					throw new Exception($"Exception while updating attachment info to Sales Order: {Company.GetLastErrorDescription()}");
			}
			return Ok("");
		}


		private Attachments2 GetAttachment(int AtcEntry)
		{
			Attachments2 atc = Company.GetBusinessObject(BaseLayer.DI.BoObjectTypes.oAttachments2);
			if (AtcEntry > 0)
			{
				if (!atc.GetByKey(AtcEntry))
				{
					throw new System.Exception($"Fatal error. Couldn't load Attachment witch AbsoluteEntry: {AtcEntry}");
				}
			}
			return atc;
		}
		
		private Documents GetSalesOrder(int DocEntry)
		{
			Documents SalesOrder = Company.GetBusinessObject(BaseLayer.DI.BoObjectTypes.oOrders);
			if (!SalesOrder.GetByKey(DocEntry))
				throw new System.Exception($"Couldn't load Sales Order with DocEntry: {DocEntry}");

			return SalesOrder;			
		}

		public string GetAttachmentsPath()
		{
			var result = "";

			var qm = new QueryManager();
			qm.CommandText = "select \"AttachPath\" from \"OADP\" ";

			using (var rs = qm.Execute(Session.Token))
			{
				result = rs.Fields.Item("AttachPath").Value.ToString();
			}

			return result;
		}


	}
}
