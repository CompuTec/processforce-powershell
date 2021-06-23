sap.ui.define([
	"computec/appengine/core/BaseController",
    "sap/ui/model/json/JSONModel",

], function(BaseController,JSONModel) {
	"use strict";

	return BaseController.extend("computec.appengine.firstPlugin.controller.MyToDo", {
		
		onInit : function (){
			BaseController.prototype.onInit.call(this);
			var oData = {
				task : {
					name : "name",
					description : "description",
					priority : "priority"
				}
			};
			var oModel = new JSONModel(oData);
			this.getView().setModel(oData)

		},


		onAdd : function (oEvent){
			var oBinding = this.getBinding();
			var oData = {
				task : {
					name : "name",
					description : "description",
					priority : "priority"
				}
			};
			oBinding.create(oData);
		},
		
		onCreate : function () {
            var oList = this.byId("todoList"), 
                oBinding = oList.getBindingContext("items"),

                oContext = oBinding.create({
                    'Code' : 10,
				'DocEntry' : 10,
				'U_TaskName' : 'By Add',
				'U_Description' : 'bla bla',
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
		}
	
    });
 });
