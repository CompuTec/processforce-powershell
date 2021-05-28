sap.ui.define([
	"computec/appengine/core/BaseController"
], function (BaseController) {
	"use strict";

	return BaseController.extend("computec.appengine.firstplugin.controller.Home", {
		onInit: function () {
			BaseController.prototype.onInit.call(this);

			this.setPageName("homePageTitle");
		},

		onTodoListPress: function (env) {
			var router = this.getRouter();
			router.navTo("todo");
		},

		onToDoPress : function (env) {
			var router = this.getRouter();
			router.navTo("mytodo");
			
		},
		
		 
	});
}); 