sap.ui.define([
	"computec/appengine/core/BaseController",
    "sap/ui/model/json/JSONModel",
	"sap/m/MessageToast",
	"sap/m/MessageBox",
	"sap/ui/core/Fragment",
	"computec/appengine/ui/model/http/Http"

], function(BaseController,
	JSONModel,
	MessageToast,
	MessageBox,
	Fragment,
	Http) {
	"use strict";

	return BaseController.extend("computec.appengine.firstPlugin.controller.MyToDo", {
		
		onInit : function (){
			BaseController.prototype.onInit.call(this);

			this.setPageName("First Plugin");

            var oViewModel = new JSONModel({
                hasUIChanges: false,
                order: 0
            });

            this.getView().setModel(oViewModel, "todoView");

		},


		onAdd : function (oEvent){
			var oBinding = this.getBinding();
			var oDatak = {
					U_TaskName : "name",
					description : "description",
					priority : "priority"
				
			};
			oBinding.create(oDatak);
		},

		onDelete: function (oEvent) {
            oEvent.getSource().getBindingContext("FP").delete("$auto").then(function () {
                MessageToast.show("deleted");
            }.bind(this), function (oError) {
                MessageBox.error(oError.message);
            });
        },
		
		onCreate : function () {
            var oList = this.byId("todoList"), 
                oBinding = oList.getBindingContext("items"),

                oContext = oBinding.create({
                'Code' : 10,
				'DocEntry' : 10,
				'U_TaskName' : 'By Add',
				'U_Description' : 'by add description',
				'U_Priority' : 'M'
                });

            this._setUIChanges(true);

            oList.getItems().some(function (oItem) {
                if (oItem.getBindingContext() === oContext) {
                    oItem.focus();
                    oItem.setSelected(true);
                    return true;
                }
            });
        },
		_setUIChanges: function (bHasUIChanges) {
            if (bHasUIChanges === undefined) {
                bHasUIChanges = this.getView().getModel().hasPendingChanges();
            }

            var oModel = this.getView().getModel("todoView");
            oModel.setProperty("/hasUIChanges", bHasUIChanges);
        },
		getBinding : function () {
			return this.getTable().getBinding("items");
		},

		getTable : function () {
			return this.byId("todoList");
		},



		onOpenDialog : function (data){
			var oView = this.getView();
			
			if(!this.byId("SalesOrderAttachment")){
				Fragment.load({
					id : oView.getId(),
					name : "computec.appengine.firstplugin.view.SalesOrderAttachment",
					controller : this
				}).then(function (oDialog){
					
					oView.addDependent(oDialog);
					oDialog.setModel(new JSONModel(data),"AT");
					oDialog.open();

				})

			}
			else{
				this.byId("SalesOrderAttachment").setModel(new JSONModel(data),"AT");
				this.byId("SalesOrderAttachment").open();
				
			}
		},

		onCloseDialog : function(){
			
			this.byId("SalesOrderAttachment").close();
		},

		

		onParamButton : function (oEvent) {
			const oSource = oEvent.getSource();
			const cardName = this.getCustomDataForElement(oSource, "CardName")
			
		},

		getCustomDataForElement: function (oElement, sCustomDataCode) {
			let oCustomData = oElement.getCustomData().find(x => x.getKey() === sCustomDataCode);
			if (oCustomData)
				return oCustomData.getValue();
			return null;
		},
		getAttachmentsByDocEntry : function(sDocNum){
			const sUrl = encodeURIComponent(`odata/CustomViews/Views.CustomWithParameters(Id='FirstPlugin:Attachments',Parameters=["AbsEntry=${sDocNum}"],paramType=Default.ParamType'Custom')`);

			return this._get(sUrl);
		},

		onParamAttachmentButton : async function (oEvent) {
			const oSource = oEvent.getSource();
			const Doc = this.getCustomDataForElement(oSource, "DocNum");
			const data = await this.getAttachmentsByDocEntry(Doc);
			this.onOpenDialog(data.value);
		},

		onDownload :  function(oEvent){
			
			const oSource = oEvent.getSource();
			const AbsEntry = this.getCustomDataForElement(oSource, "AbsEntry");
			const Line = this.getCustomDataForElement(oSource, "Line");
			const sUrl = encodeURIComponent(`http://localhost:54000/api/Attachments/GetAttachmentByCustomKey/ORDR/DocEntry/${AbsEntry}/0/${Line}`);

			return this._get(sUrl).then(response => {
				console.log(response);
			}).catch(e => {
				console.log(e);
			});
			
				
			
		},

		onDownloadInNewTab : function (oEvent){
			const oSource = oEvent.getSource();
			const AbsEntry = this.getCustomDataForElement(oSource, "AbsEntry");
			const Line = this.getCustomDataForElement(oSource, "Line");
			const sUrl = `http://localhost:54000/api/Attachments/GetAttachmentByCustomKey/ORDR/DocEntry/${AbsEntry}/0/${Line}`;
			window.open(sUrl, '_blank');
		},

		onAddAttachment : function(){
			const oFileUploader = this.getView().byId("FileUploader");
			const that = this;
			let domRef = oFileUploader.getFocusDomRef(),
					file = domRef.files[0];
			if (!file) {
				alert("No File Uploaded!");
				return;
			}
			const fromData = new FormData();
			fromData.append("file", file)
			fetch("http://localhost:54000/api/Attachments/SetAttachment/false/false" ,{
				method: 'POST',
				body: fromData
			}).then((response) => {
				console.log(response)
			})
			
			
		},

		onCountButton : async function (oEvent){
			const oSource = oEvent.getSource();
			const Name = encodeURIComponent(this.getCustomDataForElement(oSource, "CountName"));
			const sUrl = `api/FirstPlugin/Count?supplier=${Name}`;
			var p = await this._get(sUrl);
			alert(p);
		},

		_get: function (sUrl) {
			return new Promise((resolve, reject) => {
				Http.request({
					method: 'GET',
					withAuth: true,
					url: sUrl,
					done: resolve,
					fail: reject
				})
			}
			)},
			_post: function (sData, sUrl) {
				return new Promise((resolve, reject) => {
					Http.request({
						method: 'POST',
						withAuth: true,
						url: sUrl,
						data: sData,
						done: resolve,
						fail: reject
					});
				});
			},
	
    });
 });
