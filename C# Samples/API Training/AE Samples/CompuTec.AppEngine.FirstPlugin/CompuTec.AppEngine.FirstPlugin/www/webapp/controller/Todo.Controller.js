sap.ui.define([
	"computec/appengine/core/BaseController",
    "sap/ui/model/json/JSONModel",
	"sap/m/MessageToast",
	"sap/m/MessageBox",
	"computec/appengine/ui/model/http/Http",
	"sap/ui/core/Fragment"

], function(BaseController,
	JSONModel,
	MessageToast,
	MessageBox,
	Http,
	Fragment) {
	"use strict";

	return BaseController.extend("computec.appengine.firstPlugin.controller.ToDo", {
		
		onInit : function (){
			BaseController.prototype.onInit.call(this);
			this.setPageName("First Plugin");
		},
		onAdd : function (oEvent){
			var oBinding = this.getBinding();
			var oDatak = oEvent.getSource().getModel().getData();
			oBinding.create(oDatak);
		},

		onDelete: function (oEvent) {
            oEvent.getSource().getBindingContext("FP").delete("$auto").then(function () {
                MessageToast.show("deleted");
            }.bind(this), function (oError) {
                MessageBox.error(oError.message);
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

		AddFragment : async function (data){
			const that = this;
			var oView = this.getView();
			const fnOpenDialog = function (){
				const oViewModel = new JSONModel({
					U_TaskName : "n",
					U_Description : "d",
					U_Priority : "S"
				})
				that.getView().setModel(oViewModel, "model");
				that._taskDialog.open();
				
			} 
			if(!this._taskDialog){
				Fragment.load({
					id : oView.getId(),
					name : "computec.appengine.firstplugin.view.AddToDo",
					controller : this
				}).then(function (oDialog){
					
					oView.addDependent(oDialog);
					that._taskDialog = oDialog;
					fnOpenDialog();

				})

			}
			else{
				fnOpenDialog();
				
			}
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
   