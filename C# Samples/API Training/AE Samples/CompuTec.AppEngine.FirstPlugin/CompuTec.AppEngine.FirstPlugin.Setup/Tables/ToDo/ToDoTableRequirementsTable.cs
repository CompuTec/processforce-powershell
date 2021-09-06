using CompuTec.BaseLayer.Connection;
using CompuTec.BaseLayer.DI;
using CompuTec.Core2.DI.Setup.Attributes;
using CompuTec.Core2.DI.Setup.UDO.Model;
using System;
using System.Collections.Generic;

namespace CompuTec.AppEngine.FirstPlugin.Setup.Tables.ToDo
{
	[TableInstall]
	public class ToDoTableRequirementsTable : UDOManager
	{
		public const String OBJECT_CODE = "SAMPLE_TO_DO_REQ";
		public const String TABLE_NAME = "SAMPLE_TDO1";
		public const String TABLE_DESCRIPTION = "Sample To Do Requirements";
		public const String ARCHIVE_TABLE_NAME = "SAMPLE_ATDO1";

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
			var ItemName = new UDOTableField();
			ItemName.SetName("ItemName");
			ItemName.SetDescription("Item Name");
			ItemName.SetType(BoFieldTypes.db_Alpha);
			ItemName.SetEditSize(100);
			fields.Add(ItemName);

			//description column
			var Quantity = new UDOTableField();
			Quantity.SetName("Quantity");
			Quantity.SetDescription("Quantity");
			Quantity.SetType(BoFieldTypes.db_Numeric);
			Quantity.SetEditSize(11);
			fields.Add(Quantity);

			return fields;
		}

		protected override void SetChildTables()
		{
		}

	}
}
