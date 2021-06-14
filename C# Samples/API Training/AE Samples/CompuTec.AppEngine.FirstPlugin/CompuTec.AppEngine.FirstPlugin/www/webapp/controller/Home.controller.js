sap.ui.define([
	"computec/appengine/core/BaseController"
], function (BaseController) {
	"use strict";

	return BaseController.extend("computec.appengine.firstplugin.controller.Home", {
		onInit: function () {
			BaseController.prototype.onInit.call(this);

			this.setPageName("homePageTitle");
		},

		onSalesOrderPress: function (env) {
			var router = this.getRouter();
			router.navTo("salesorder");
		},

		onToDoPress : function (env) {
			var router = this.getRouter();
			router.navTo("todo");
			
		},
		
		 
	});
});   