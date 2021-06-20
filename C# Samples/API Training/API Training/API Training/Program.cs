using CompuTec.Core.DI.Database;
using CompuTec.ProcessForce.API;
using CompuTec.ProcessForce.API.Core;
using CompuTec.ProcessForce.API.Documents.BillOfMaterials;
using CompuTec.ProcessForce.API.Documents.ItemDetails;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace API_Training
{
    class Program
    {
        private static IProcessForceCompany pfCompany;

        public static void UseSApObject()
        {
            SAPbobsCOM.Items itm = pfCompany.SapCompany.GetBusinessObject(BoObjectTypes.oItems);
            itm.ItemCode = "testItm";
            itm.Add();
        }
        /// <summary>
        /// One Line Query
        /// </summary>
        public static void UseQueries1()
        {
            //select "ItemName" from OITM where "ItemCode"='Active-Item-01' 
            var itemName= QueryManager.ExecuteSimpleQuery<string>(pfCompany.Token, "OITM", "ItemName", "ItemCode", "Active-Item-01");
        }
        //Simple queries
        public static void UseQueries2()
        {
            //select "ItemName","ItmGrpCode" from OITM where "ItemCode" in ('Active-Item-01', 'Active-Item-02')

            var qm = new QueryManager();
            qm.SimpleTableName = "OITM";
            qm.SetSimpleWhereFields("ItemCode");
            qm.SetSimpleResultFields("ItemName", "ItmGrpCod");
            using (var rs = qm.ExecuteSimpleParameters(pfCompany.Token, QueryManager.GenerateSqlInStatment<string>(new string[] { "Active-Item-01", "Active-Item-02" })))
            {
                for (int i = 0; i < rs.RecordCount; i++)
                {
                    var itmname = rs.Fields.Item("ItemName").Value;
                    var itmGrpCode= rs.Fields.Item("ItmGrpCod").Value;
                    rs.MoveNext();
                }
            }

        }
        const string Query = @" select t0.""ItemCode"" from ""OITM"" t0 inner join ""OITW"" t1 on t0.""ItemCode""=t1.""ItemCode"" where t1.""WhsCode"" like @param1 and t0.""ItmGrpCod"" = @GrpCode";
        //most complicated queries
        public static void UseQueries3()
        {
            var qm = new QueryManager();
            qm.CommandText = Query;
            qm.AddParameter("param1", "a$");
            qm.AddParameter("GrpCode", 12);
            using (var rs = qm.Execute(pfCompany.Token))
            {
                for (int i = 0; i < rs.RecordCount; i++)
                {
                    var itmCode = rs.Fields.Item(0).Value;
                    rs.MoveNext();
                }
            }
        }
        static void Main(string[] args)
        {
            Connect();

           var entry= CreateManufacturingOrder("Product-A", "code00", 100);
          //var reslt=  pfCompany.GenerateAdditionalBatchDetails(0, 0);
          //  if (reslt.Success)
          //      Console.WriteLine("OK");
          //  else
          //      Console.WriteLine("NOP");
          // var batchinformations=CompuTec.ProcessForce.API.Generators.BatchNumberGenerator.GenerateBatch(pfCompany.Token, "A", "","revision00","default",true);
   
          //  CreateProductionGoodsReceipt(479);
          //// UpdatePickReceipt(479);
          //// CreateProductionGoodsIssue(481);
          // //UpdatePickOrder(481);
          // //RestoreItemDetails()
          //  var ItemCode="fg1";
          // string Rev = string.Empty;
          // CreateNewRevisionInItemDetails(ItemCode, out Rev);
          // pfCompany.SapCompany.StartTransaction();
          // try
          // {
          //     CreateNewBillOfMaterial(ItemCode, Rev);
          //     var mor = CreateManufacturingOrder(ItemCode, Rev, 10d);
          //     var por = CreatePickOrder(mor);
          //     var pre = CreatePickReceipt(mor);
          //     UpdatePickOrder(por);
          //     UpdatePickReceipt(pre);
          //     CreateProductionGoodsIssue(por);
          //     CreateProductionGoodsReceipt(pre);
          //     pfCompany.SapCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
          // }
          // catch
          // {
          //     if(pfCompany.SapCompany.InTransaction)
          //         pfCompany.SapCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);

            // }
        }
        private static void Connect()
        {
            
            pfCompany = ProcessForceCompanyInitializator.CreateCompany();
            try
            {
               // CompuTec.Core.Connection.ConnectionHolder.ConType = CompuTec.Core.Connection.ConnectionType.DI;
                //set connection constaint
                pfCompany.SLDAddress = "hanadev:40000";
                pfCompany.UserName = "manager";
                pfCompany.Password = "1234";
                pfCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB;
                pfCompany.Databasename = "PROD_20210122_SPROC_11212";
                pfCompany.Language = SAPbobsCOM.BoSuppLangs.ln_English;
                //Connect to Company
                if (pfCompany.Connect() == 1)
                    Console.WriteLine(string.Format("Connected to Company {0}", pfCompany.SapCompany.CompanyName));
                else
                    Console.WriteLine(string.Format("Not Connected To compmany Error:{0}", pfCompany.getLastErrorDescription()));
            }
            catch (Exception ex)
            {
                Console.WriteLine(string.Format("Exception was throw {0}", ex.Message));
            }
            finally
            {

                //Console.ReadKey();
                //if (pfCompany.IsConnected)
                //    pfCompany.Disconnect();
            } 
        }

        /// <summary>
        /// You must restore Item Details after adding new Item Master Data from SDK or B1IF
        /// </summary>
        private static void RestoreItemDetails()
        {
            Recordset rec = pfCompany.CreateSapObject(BoObjectTypes.BoRecordset) as Recordset;
            rec.DoQuery(" select t0.ItemCode from OITM t0 left outer join [@CT_PF_OIDT] t1 on t0.ItemCode=t1.U_ItemCode where t1.Code is null");
            if (rec.RecordCount == 0)
                return;
            var Fields=rec.Fields;
            var itemCodeValue = Fields.Item(0);
            try
            {
                for (int i = 0; i < rec.RecordCount; i++)
                {
                    IItemDetails idt = pfCompany.CreatePFObject(ObjectTypes.ItemDetails);
                    idt.U_ItemCode = itemCodeValue.Value;
                    idt.Add();
                }
            }
            finally
            {
                Marshal.ReleaseComObject(itemCodeValue);
                Marshal.ReleaseComObject(Fields);
                Marshal.ReleaseComObject(rec);
            }
        }
        private static void CreateNewRevisionInItemDetails(string ItemCode, out string  RevisionCode )
        {
            IItemDetails idt = pfCompany.CreatePFObject(ObjectTypes.ItemDetails);
              idt.GetByItemCode(ItemCode);
              var revisionCount = idt.Revisions.Count;
              idt.Revisions.SetCurrentLine(revisionCount - 1);
            if(!string.IsNullOrEmpty(idt.Revisions.U_Code))
            {
                //Add empty Line
                revisionCount++;
                idt.Revisions.Add();
            }
            RevisionCode=string.Format("Rev0{0}", revisionCount);
            idt.Revisions.U_Code = string.Format("Rev0{0}", revisionCount);
            idt.Revisions.U_Description = string.Format("Revision No {0}", revisionCount);
            idt.Revisions.U_Status = RevisionStatus.Active;
            if (idt.Update() == 0)
            {
                Console.WriteLine(string.Format(" Revision {0} for Item{0} was added", string.Format("Rev0{0}", revisionCount), ItemCode));
            }
            else
            {
                Console.WriteLine(string.Format(" Revision {0} for Item{0} was not added", string.Format("Rev0{0}", revisionCount), ItemCode));
            }
        }

        private static void CreateNewBillOfMaterial(string ItemCOde, string Revision)
        {
            IBillOfMaterial bom = pfCompany.CreatePFObject(CompuTec.ProcessForce.API.Core.ObjectTypes.BillOfMaterial);
            bom.U_ItemCode = ItemCOde;
            bom.U_Revision = Revision;
            bom.Items.U_ItemCode = "Active-Item-01";
            bom.Items.U_WhsCode = "01";
            bom.Items.U_Quantity = 0.5d;
            bom.Items.U_IssueType = "B";
            bom.Items.Add();
            bom.Items.U_ItemCode = "Active-Item-02";
            bom.Items.U_Quantity = 0.75d;
            bom.Items.U_WhsCode = "01";
            bom.Items.U_IssueType = "M";
            bom.Routings.U_RtgCode = "01";
            bom.Routings.U_IsDefault = "Y";
            bom.Routings.U_IsRollUpDefault = "Y";
            bom.Add();
        }
        public static int  CreateManufacturingOrder (string ItemCode,string Revision, double qty)
        {
            CompuTec.ProcessForce.API.Documents.ManufacturingOrder.IManufacturingOrder mor = pfCompany.CreatePFObject(ObjectTypes.ManufacturingOrder);
            var BomCode = GetBomCode(ItemCode, Revision);
            mor.U_BOMCode = BomCode;

            //mor.U_ItemCode = "ItemA";
            //mor.U_Revision = "revision01";
            mor.U_Quantity = qty;
            mor.U_RequiredDate = DateTime.Today.AddDays(5);
            mor.U_SchedulingMtd = CompuTec.ProcessForce.API.Documents.ManufacturingOrder.PF_MORSchedulingMthd.Forward;
            //Schedule Mor
            mor.CalculateManufacturingTimes(false);
            mor.U_Status = CompuTec.ProcessForce.API.Enumerators.ManufacturingOrderStatus.Released;
   
            if (mor.Add() == 0)
            {
                return Convert.ToInt32(mor.GetLastObjectCode());
            }
            else
            {
                Console.WriteLine(string.Format("Cannot Add MOR {0}", ItemCode));
                return -1;
            }
        }

        public static int CreatePickOrder(int DocEntry)
        {
            object porEntry = 0;
            CompuTec.ProcessForce.API.Actions.CreatePickOrderForProductionIssue.ICreatePickOrderForProductionIssue action = pfCompany.CreatePFAction(ActionType.CreatePickOrderForProductionIssue);
            action.AddMORDocEntry(DocEntry);
            action.DoAction(out porEntry);
            return Convert.ToInt32( porEntry);

        }
        public static int CreatePickReceipt(int Entry)
        {
            object porEntry = 0;
            CompuTec.ProcessForce.API.Actions.CreatePickReceiptForProductionReceipt.ICreatePickReceiptForProductionReceipt action = pfCompany.CreatePFAction(ActionType.CreatePickReceiptForProductionReceipt);
            action.AddManufacturingOrderDocEntry(Entry);
            action.DoAction(out porEntry);
            return Convert.ToInt32(porEntry);
        }
        private static string GetBomCode(string ItemCode, string Revision)
        {
             Recordset rec = pfCompany.CreateSapObject(BoObjectTypes.BoRecordset) as Recordset;
            rec.DoQuery(string.Format(" select Code from [@CT_PF_OBOM] where U_ItemCode =N'{0}' and U_Revision=N'{1}'",ItemCode,Revision));
            if (rec.RecordCount == 0)
            {
                Console.WriteLine(string.Format("cannot Find BOM for ID={0} and Rev={1}", ItemCode, Revision));
                return string.Empty;
                Marshal.ReleaseComObject(rec);
            }
            var rcode = rec.Fields.Item(0).Value;
            Marshal.ReleaseComObject(rec);
            return rcode;
        }
        public static void UpdatePickReceipt(int de)
        {
            CompuTec.ProcessForce.API.Documents.PickReceipt.IPickReceipt pre = pfCompany.CreatePFObject(ObjectTypes.PickReceipt);
            pre.GetByKey(de.ToString());
            pre.RequiredItems.SetCurrentLine(0);
            pre.RequiredItems.U_PickedQty = 5;
            var reqitemsLineNum = pre.RequiredItems.U_LineNum;
            pre.PickedItems.SetCurrentLine(0);
            #region Foreach Batch to pick
            pre.PickedItems.U_ItemCode = pre.RequiredItems.U_ItemCode;
            pre.PickedItems.U_BnDistNumber = pfCompany.GenerateBatchNumber(pre.PickedItems.U_ItemCode, pre.RequiredItems.U_Classification);
            pre.PickedItems.U_Quantity = 5;
            pre.PickedItems.U_ReqItmLn = reqitemsLineNum;
            var pickedItemsLineNum = pre.PickedItems.U_LineNum;
            //pre.Relations.U_PickItemLineNo = pickedItemsLineNum; // relations are obsolete since 9.2 PL09 version - use U_ReqItmLine instead (as second line above)
            //pre.Relations.U_ReqItemLineNo = reqitemsLineNum;
            //pre.Relations.Add();
            #region If Bin Allocation Enabled on Whs

            pre.BinAllocations.SetCurrentLine(0);
            pre.BinAllocations.U_SnAndBnLine = pickedItemsLineNum;
            pre.BinAllocations.U_BinAbsEntry = 2;
            pre.BinAllocations.U_Quantity = 5;
            pre.BinAllocations.Add();
            #endregion
            pre.PickedItems.Add(); 
            #endregion
            pre.Update();
        }
        public static void UpdatePickOrder(int DE)
        {
            CompuTec.ProcessForce.API.Documents.PickOrder.IPickOrder por = pfCompany.CreatePFObject(ObjectTypes.PickOrder);
            por.GetByKey(DE.ToString());
            por.RequiredItems.SetCurrentLine(0);
            var reqItemLineNUm = por.RequiredItems.U_LineNum;
            por.RequiredItems.U_PickedQty = 1;
            por.PickedItems.SetCurrentLine(0);


            #region Foreach Batch to pick
            var pickedItemsLineNum = por.PickedItems.U_LineNum;
            por.PickedItems.U_ItemCode = por.RequiredItems.U_ItemCode;
            por.PickedItems.U_BnDistNumber = "2012-04-18-3";
            por.PickedItems.U_Quantity = 1;
            por.PickedItems.U_ReqItmLn = por.RequiredItems.U_LineNum;
            //por.Relations.U_PickItemLineNo = pickedItemsLineNum; // relations are obsolete since 9.2 PL09 version - use U_ReqItmLine instead (as line above)
            //por.Relations.U_ReqItemLineNo = reqItemLineNUm;
            //por.Relations.Add();
            #region  IF Bin Allocation on whs enabled
            por.BinAllocations.U_BinAbsEntry = 2;
            por.BinAllocations.U_Quantity = 1;
            por.BinAllocations.U_SnAndBnLine = pickedItemsLineNum;
            por.BinAllocations.Add(); 
            #endregion
            por.PickedItems.Add(); 
            #endregion
    
            por.Update();
        }
        public static void CreateProductionGoodsIssue(int docentry)
        {
            CompuTec.ProcessForce.API.Actions.CreateGoodsIssueFromPickOrderBasedOnProductionIssue.ICreateGoodsIssueFromPickOrderBasedOnProductionIssue action = pfCompany.CreatePFAction(ActionType.CreateGoodsIssueFromPickOrderBasedOnProductionIssue);
            action.PickOrderID = docentry;
            action.DocDate = DateTime.Today;
            action.TaxDate = DateTime.Today; 
            object outnumber=0;
            action.DoAction(out outnumber);
        }
        public static void CreateProductionGoodsReceipt(int docEntry)
            
        {
            CompuTec.ProcessForce.API.Actions.CreateGoodsReceiptFromPickReceiptBasedOnProductionReceipt.ICreateGoodsReceiptFromPickReceiptBasedOnProductionReceipt action = pfCompany.CreatePFAction(ActionType.CreateGoodsReceiptFromPickReceiptBasedOnProductionReceipt);
            action.DocDate = DateTime.Today;
            action.TaxDate = DateTime.Today;
            action.PickReceiptID = docEntry;
            //Must Be filled 
            action.DocumentMemo = "Created By API";
            object docNumber=0;
            action.DoAction(out docNumber);

        }
    }
}
