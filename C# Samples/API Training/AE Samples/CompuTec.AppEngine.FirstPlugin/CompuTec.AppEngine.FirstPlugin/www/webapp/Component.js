// @ts-ignore
jQuery.sap.require("computec.appengine.plugin.PluginRouter");

sap.ui.define([
	"computec/appengine/ui/core/UIComponent"
], function (UIComponent) {
	"use strict";

	return UIComponent.extend("computec.appengine.firstplugin.Component", {

		metadata: {
			manifest: "json"
		},

		init: function () {
			let self = this;
			UIComponent.prototype.onInit.call(self);
			UIComponent.prototype.init.apply(this, arguments);

			this.getRouter().initialize();

			this.attachSLOdataModel("odata/FirstPlugin/", "FP");
		},

		destroy: function () {
			UIComponent.prototype.destroy.apply(this, arguments);
		}
	});
});