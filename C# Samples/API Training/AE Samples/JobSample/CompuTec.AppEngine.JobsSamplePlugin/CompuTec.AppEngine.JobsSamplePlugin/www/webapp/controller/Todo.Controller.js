sap.ui.define([
    "computec/appengine/core/BaseController",
    "sap/ui/model/json/JSONModel",
    "sap/m/MessageBox",
    "sap/ui/model/Sorter",
    "sap/ui/model/Filter",
    "sap/ui/model/FilterOperator",
    "sap/ui/model/FilterType",
], function (BaseController, JSONModel, MessageBox, Sorter, Filter, FilterOperator, FilterType) {
    "use strict";

    return BaseController.extend("computec.appengine.jobssampleplugin.controller.Todo", {
        onInit: function () {
            BaseController.prototype.onInit.call(this);

            this.setPageName("todoPageTitle");

            var oViewModel = new JSONModel({
                hasUIChanges: false,
                order: 0
            });

            this.getView().setModel(oViewModel, "todoView");
        },



        onCreate: function () {
            var oList = this.byId("todoList"),
                oBinding = oList.getBinding("items"),

                oContext = oBinding.create({
                    "IsDone": false,
                    "Title": ""
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

        onSave: function () {
            var fnSuccess = function () {
                this._setUIChanges(false);
            }.bind(this);

            var fnError = function (oError) {
                this._setUIChanges(false);
                MessageBox.error(oError.message);
            }.bind(this);

            this.getView().getModel("CompuTec.AppEngine.JobsSamplePlugin").submitBatch("todoGroup").then(fnSuccess, fnError);
        },

        onDelete: function (oEvent) {
            oEvent.getSource().getBindingContext("CompuTec.AppEngine.JobsSamplePlugin").delete("$auto").then(function () {
                MessageToast.show(this.geti18n().getText("deletionSuccessMessage"));
            }.bind(this), function (oError) {
                MessageBox.error(oError.message);
            });
        },



        onResetChanges: function () {
            this.byId("todoList").getBinding("items").resetChanges();
            this._setUIChanges(false);
        },

        onRefresh: function () {
            var oBinding = this.byId("todoList").getBinding("items");

            if (oBinding.hasPendingChanges()) {
                this.byId("todoList").getBinding("items").resetChanges();
            }

            oBinding.refresh();
        },

        onInputChange: function (oEvent) {
            if (oEvent.getParameter("escPressed")) {
                this._setUIChanges(fasle);
            } else {
                this._setUIChanges(true);
            }
        },


        onSearch: function () {
            var oView = this.getView(),
                sValue = oView.byId("searchField").getValue(),
                oFilter = new Filter("Title", FilterOperator.Contains, sValue);

            oView.byId("todoList").getBinding("items").filter(oFilter, FilterType.Application);
        },


        _setUIChanges: function (bHasUIChanges) {
            if (bHasUIChanges === undefined) {
                bHasUIChanges = this.getView().getModel().hasPendingChanges();
            }

            var oModel = this.getView().getModel("todoView");
            oModel.setProperty("/hasUIChanges", bHasUIChanges);
        },
    });
});