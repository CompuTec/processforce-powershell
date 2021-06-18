using CompuTec.BaseLayer.Connection;
using CompuTec.BaseLayer.DI;
using CompuTec.Core2.Beans.DataLayer.UDOXmlStructure;
using CompuTec.Core2.DI.Setup.Attributes;
using CompuTec.Core2.DI.Setup.UDO.Model;
using System;
using System.Collections.Generic;


namespace CompuTec.AppEngine.First.DBInstall.Tables
{
	[TableInstall]
	public class ToDoTable : UDOManager
	{
		public const String OBJECT_CODE = "Sample_ToDo";
		public const String TABLE_NAME = "SAMPLE_TODO";
		public const String TABLE_DESCRIPTION = "Sample table";
		public const String ARCHIVE_TABLE_NAME = "SAMPLE_ATDO";

		public ToDoTable(IDIConnection connection) : base(connection) { }


        protected override IUDOTable CreateUDOTable()
		{
			List<IUDOField> fields = this.CreateFieldsForHeaderTable();
			List<IUDOFindColumn> findColumns = this.CreateFindColumnsList();

			IUDOTable UdoTable = new UDOTable(fields, findColumns, TABLE_NAME, TABLE_DESCRIPTION, BoUTBTableType.bott_MasterData, this.CreateKeys());

			UdoTable.RegisteredUDOName = TABLE_NAME;
			UdoTable.RegisteredUDOCode = OBJECT_CODE;

			UdoTable.CanArchive = BoYesNoEnum.tYES;
			UdoTable.CanCancel = BoYesNoEnum.tNO;
			UdoTable.CanClose = BoYesNoEnum.tNO;
			UdoTable.CanCreateDefaultForm = BoYesNoEnum.tNO;
			UdoTable.CanDelete = BoYesNoEnum.tYES;
			UdoTable.CanFind = BoYesNoEnum.tYES;
			UdoTable.CanLog = BoYesNoEnum.tYES;
			UdoTable.CanYearTransfer = BoYesNoEnum.tNO;
			UdoTable.ArchiveTableName = ARCHIVE_TABLE_NAME;

			return UdoTable;
		}

		private List<IUDOFindColumn> CreateFindColumnsList()
		{
			List<IUDOFindColumn> findList = new List<IUDOFindColumn>();

			var taskName = new UDOFindColumn();
			taskName.SetColumnAlias("U_TaskName");
			taskName.SetColumnDescription("Task Name");
			findList.Add(taskName);

			return findList;
		}

		private List<IUDOField> CreateFieldsForHeaderTable()
		{
			var fields = new List<IUDOField>();

			//adding task name column
			var TaskName = new UDOTableField();
			TaskName.SetName("TaskName");
			TaskName.SetDescription("Task Name");
			TaskName.SetType(BoFieldTypes.db_Alpha);
			TaskName.SetEditSize(100);
			fields.Add(TaskName);


			//description column
			var TaskDescription = new UDOTableField();
			TaskDescription.SetName("Description");
			TaskDescription.SetDescription("Task description");
			TaskDescription.SetType(BoFieldTypes.db_Alpha);
			TaskDescription.SetEditSize(254);
			fields.Add(TaskDescription);

			//priority column
			var TaskPriority = new UDOTableField();
			TaskPriority.SetName("Priority");
			TaskPriority.SetDescription("priority of the task");
			TaskPriority.SetType(BoFieldTypes.db_Alpha);
			TaskPriority.ValidValuesMD = new Dictionary<string, string>()
			{
				{ "L","Low Priority" },
				{ "M", "Medium Priority" },
				{ "H", "Huge Priority" }
			};
			TaskPriority.DefaultValue = "L";
			TaskPriority.SetEditSize(1);
			fields.Add(TaskPriority);


			//deadline column
			var TaskDeadline = new UDOTableField();
			TaskDeadline.SetName("Deadline");
			TaskDeadline.SetDescription("Deadline");
			TaskDeadline.SetType(BoFieldTypes.db_Date);
			TaskDeadline.SetEditSize(10);
			fields.Add(TaskDeadline);


			return fields;

		}
		private List<IUDOTableKey> CreateKeys()
		{
			List<IUDOTableKey> list = new List<IUDOTableKey>();
			return list;
		}

		protected override void SetChildTables()
		{
			ChildTablesClasses.AddRange(new string[] { "RequirementTable" });
		}
	}
}