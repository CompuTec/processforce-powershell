﻿using CompuTec.BaseLayer.Connection;
using CompuTec.BaseLayer.DI;
using CompuTec.Core2.Beans.DataLayer.UDOXmlStructure;
using CompuTec.Core2.DI.Setup.Attributes;
using CompuTec.Core2.DI.Setup.UDO.Model;
using System;
using System.Collections.Generic;


namespace CompuTec.AppEngine.FirstPlugin.Setup.DBInstall.Tables.ToDoObjectDefinition
{
	[TableInstall]
	public class ToDoTableRequirementsTable : UDOManager
	{
		public const String OBJECT_CODE = "ToDoRequirement";
		public const String TABLE_NAME = "CT_TST_TDO1";
		public const String TABLE_DESCRIPTION = "ToDo:Requirementstable";
		public const String ARCHIVE_TABLE_NAME = "CT_TST_ATDO1";

		public ToDoTableRequirementsTable(IDIConnection connection) : base(connection) { }


		protected override IUDOTable CreateUDOTable()
		{
			List<IUDOField> fields = this.DefineChildFields();

			IUDOTable UdoTable = new UDOTable(fields, TABLE_NAME, TABLE_DESCRIPTION, BoUTBTableType.bott_MasterDataLines);

			UdoTable.RegisteredUDOName = TABLE_NAME;
			UdoTable.RegisteredUDOCode = OBJECT_CODE;
			UdoTable.ArchiveTableName = ARCHIVE_TABLE_NAME;

			return UdoTable;
		}

		private List<IUDOField> DefineChildFields()
		{
			var fields = new List<IUDOField>();

			//adding task name column
			var TaskName = new UDOTableField();
			TaskName.SetName("Name");
			TaskName.SetDescription("Requirement Name");
			TaskName.SetType(BoFieldTypes.db_Alpha);
			TaskName.SetEditSize(100);
			fields.Add(TaskName);


			//description column
			var TaskDescription = new UDOTableField();
			TaskDescription.SetName("Quantity");
			TaskDescription.SetDescription("Quantity");
			TaskDescription.SetType(BoFieldTypes.db_Numeric);
			TaskDescription.SetEditSize(11);
			fields.Add(TaskDescription);

			

			return fields;

		}

		protected override void SetChildTables()
		{
		}

	}
}