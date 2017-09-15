
import ViewExport = require("scripts/viewExport");
import ViewImport = require("scripts/viewImport");

var _viewImport: ViewImport.viewImport;
var _viewExport: ViewExport.viewExport;
_viewImport = new ViewImport.viewImport();
_viewImport.init();

_viewExport = new ViewExport.viewExport();
_viewExport.init();

export function registerTab() {

  

    $("#selectAction").on("change", e => {
        $("#tabExport").hide();
        $("#tabImport").hide();
        $($("#selectAction").val()).show();
    });

    var contrib = "ote-import-tab";//VSS.getContribution().id;
 //   contrib =VSS.getContribution().id;


    updateContext(VSS.getConfiguration());
    console.log("Registered " + contrib);
    VSS.register(contrib, {

        pageTitle: function (state) {
            console.log("page title:state ");
            console.log(state);

            return "Test suite: " + state.selectedSuite.name + "(Suite ID: " + state.selectedSuite.id +")";

        },
        updateContext: tabContext => {
            updateContext(tabContext);
        },
        isInvisible: function (state) {
            // Hide this tab if the user has selected the "Unsaved work items" pseudo-query.
            return false;
        }
    });

}

function updateContext(tabContext) {
    console.log("updateContext");
    _viewExport.updateContext(tabContext);
    _viewImport.updateContext(tabContext);
    console.log(tabContext);
}