using CompuTec.AppEngine.Base.Infrastructure.Jobs;
using CompuTec.AppEngine.Base.Infrastructure.Jobs.Annotations;
using CompuTec.AppEngine.Base.Infrastructure.Security;
using CompuTec.AppEngine.FirstPlugin.API.BusinessObjects.ToDo;
using CompuTec.BaseLayer.DI;
using CompuTec.Core2.DI.Database;
using Newtonsoft.Json;
using NLog;
using StructureMap;
using System;

namespace CompuTec.AppEngine.FirstPlugin.Jobs
{
	[EventBusJob(JobId = "SalesOrderToApproveEventJob", Description = "Crate new To Do Job for Added Sales Orders that are unapproved", ContentType = "17", ActionType = "A")]
	public class SalesOrderToApproveEventJob : EventBusSecureJob
	{
		Logger _logger;
		public SalesOrderToApproveEventJob(Session session, IContainer container, EventBus.Message message) : base(session, container, message)
		{
			_logger = container.GetInstance<Logger>();
		}

		public override void Call()
		{
			try
			{
				_logger.Trace($"Job :SalesOrderToApproveEventJob Started for :{Message.Body}");
				dynamic json = JsonConvert.DeserializeObject(Message.Body);
				int DocEntry = json.DocEntry;

				bool approved;
				int DocNum;
				using (CTRecordset rs = this.GetSalesOrderDetails(DocEntry))
				{
					DocNum = rs.Fields.Item("DocNum").Value;
					string Confirmed = rs.Fields.Item("Confirmed").Value;
					approved = Confirmed == "Y" ? true : false;
				}

				if (!approved)
				{
					AddNewToDoTask(DocNum);
					_logger.Trace($"Job :SalesOrderToApproveEventJob finished successfully. To Do Task added");
				}
				else
				{
					_logger.Trace($"Job :SalesOrderToApproveEventJob finished successfully - Nothing to do, Sales Order already approved.");
				}
			}
			catch (Exception e)
			{
				_logger.Error(e, $"Job :SalesOrderToApproveEventJob failed:{e.Message}");
				throw;
			}

		}
		private void AddNewToDoTask(int DocNum)
		{

			IToDo toDoTask = CompuTec.Core2.CoreManager.GetUDO(Session.Token, "SAMPLE_TO_DO");
			toDoTask.U_TaskName = $"Confirmation";
			toDoTask.U_Description = $"Review Sales Order number {DocNum}";
			toDoTask.U_Priority = API.Enums.ToDoPriority.Medium;
			if (toDoTask.Add() != 0)
				throw new Exception($"Exception while adding ToDo task: {Session.Company.GetLastErrorDescription()}");
		}

		private CTRecordset GetSalesOrderDetails(int DocEntry)
		{
			var qm = new QueryManager();
			qm.SetSimpleResultFields("DocEntry", "DocNum", "CardCode", "Confirmed");
			qm.SimpleTableName = "ORDR";
			qm.SetSimpleWhereFields("DocEntry");

			return qm.ExecuteSimpleParameters(Session.Token, DocEntry);
		}
	}

}
