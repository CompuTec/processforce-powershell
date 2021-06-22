This sample demonstrate the way to implement AE Jobs in AppEngine
1) Scheduling Job -
	this  job is executed on time based manner, you must specify CRON expression that represents how often this job will be executed
2) Event Based Job-
	this job is executed like transaction notification adter the object is added/updated/removed in SAP system 
3) OneTimeJob-
this job is executed once when AE is restarted or on demand ( when button Run is executed)


 class CompuTec.AppEngine.JobsSamplePlugin.Jobs.JobTypes.ScheduledJob
 this job uses Core2 Connection that is established in AE/Job settings  and creates an Item master Data object;
 
 class CompuTec.AppEngine.JobsSamplePlugin.Jobs.JobTypes.EventBasedJob 
 this job reacts on ItemMasterData Add and puts a comment in the item master data 

  class CompuTec.AppEngine.JobsSamplePlugin.Jobs.JobTypes.OneTimeJob 
  create a Business Partner

  class CompuTec.AppEngine.JobsSamplePlugin.Jobs.JobTypes.ProcessForceConnection
  this is Event based job, uses ProcessForce Connection when sales order is created system automatically creates Manufacturing order for the content of created document

