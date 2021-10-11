sap.ui.define([
	"computec/appengine/core/BaseController"
], function (BaseController) {
	"use strict";

	return BaseController.extend("computec.appengine.firstplugin.controller.App", {
		onInit: function () {
			BaseController.prototype.onInit.call(this);
		}
	});
});