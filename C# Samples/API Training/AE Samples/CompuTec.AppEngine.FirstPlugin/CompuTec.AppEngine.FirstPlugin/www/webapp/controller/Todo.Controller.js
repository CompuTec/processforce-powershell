sap.ui.define([
	"computec/appengine/core/BaseController",
	"sap/ui/model/json/JSONModel",
	"sap/m/MessageBox",
	"sap/m/MessageToast",
	"sap/ui/model/Sorter",
	"sap/ui/model/Filter",
	"sap/ui/model/FilterOperator",
	"sap/ui/model/FilterType",
	"computec/appengine/ui/controls/YesNoBoolType"
], function (BaseController, JSONModel, MessageBox, MessageToast,
	Sorter, Filter, FilterOperator, FilterType, YesNoBoolType) {
	"use strict";

	return BaseController.extend("computec.appengine.firstplugin.controller.Todo", {
		onInit: function () {
			BaseController.prototype.onInit.call(this);

			this.setPageName("todoPageTitle");
		},

		onCreate: function () {
			var oList = this.byId("todoList"),
				oBinding = oList.getBinding("items"),

				oContext = oBinding.create({
					U_Priority: 'Low',
					U_Done: 'No'
				});

			oList.getItems().some(function (oItem) {
				if (oItem.getBindingContext() === oContext) {
					oItem.focus();
					return true;
				}
			});
		},

		onDelete: function (oEvent) {
			oEvent.getSource().getBindingContext("FirstPlugin").delete("$auto").then(function () {
				MessageToast.show(this.geti18n().getText("deletionSuccessMessage"));
			}.bind(this), function (oError) {
				MessageBox.error(oError.message);
			});
		},


		onResetChanges: function () {
			this.byId("todoList").getBinding("items").resetChanges();
		},

		onRefresh: function () {
			var oBinding = this.byId("todoList").getBinding("items");

			if (oBinding.hasPendingChanges()) {
				this.byId("todoList").getBinding("items").resetChanges();
			}

			oBinding.refresh();
		},

		onSearch: function () {
			var oView = this.getView(),
				sValue = oView.byId("searchField").getValue();

			const aFilters = [];
			if (typeof sValue === 'string' && sValue.length > 0) {
				aFilters.push(new Filter({
					path: 'U_TaskName',
					operator: FilterOperator.Contains,
					value1: sValue
				}));
				aFilters.push(new Filter({
					path: 'U_Description',
					operator: FilterOperator.Contains,
					value1: sValue
				}));
			}

			const oFilter = new Filter({
				filters: aFilters,
				and: false
			});

			oView.byId("todoList").getBinding("items").filter(oFilter);
		},

		getIconForPriority: function (sPriority) {
			let sIcon;
			switch (sPriority) {
				case 'Low':
					break;
				case 'Medium':
					sIcon = 'sap-icon://hint';
					break;
				case 'Huge':
					sIcon = 'sap-icon://warning';
					break;
				default:
					sIcon = null;
					break;
			}
			return sIcon;
		},
		getStateForPriority: function (sPriority) {
			let sState;
			switch (sPriority) {
				case 'Low':
					break;
				case 'Medium':
					sState = 'Information';
					break;
				case 'Huge':
					sState = 'Warning';
					break;
				default:
					sState = 'None';
					break;
			}
			return sState;
		}
	});
});