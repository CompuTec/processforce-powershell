sap.ui.define([
	"computec/appengine/core/BaseController",
    "sap/ui/model/json/JSONModel",
	"sap/m/MessageToast",
	"sap/m/MessageBox",
	"sap/ui/core/Fragment"

], function(BaseController,
	JSONModel,
	MessageToast,
	MessageBox,
	Fragment) {
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
			var oModel = new JSONModel(oDatak)
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
		onOpenDialog : function (){
			var oView = this.getView();

			if(!this.byId("SalesOrderAttachment")){
				Fragment.load({
					id : oView.getId(),
					name : "computec.appengine.firstplugin.view.SalesOrderAttachment"
				}).then(function (oDialog){
					oView.addDependent(oDialog);
					oDialog.open();
				})

			}
			else{
				this.byId("SalesOrderAttachment").open();
			}
		},

		

		onParamButton : function (oEvent) {
			const oSource = oEvent.getSource();
			const cardName = this.getCustomDataForElement(oSource, "CardName")
			alert(cardName);
		},

		getCustomDataForElement: function (oElement, sCustomDataCode) {
			let oCustomData = oElement.getCustomData().find(x => x.getKey() === sCustomDataCode);
			if (oCustomData)
				return oCustomData.getValue();
			return null;
		}

	
    });
 });
