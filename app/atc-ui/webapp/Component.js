sap.ui.define(["sap/ui/core/UIComponent"], (UIComponent) => {
  "use strict";
  return UIComponent.extend("com.sap.atcanalysis.Component", {
    metadata: { manifest: "json" },
    init() {
      UIComponent.prototype.init.apply(this, arguments);
      this.getRouter().initialize();
    },
  });
});
