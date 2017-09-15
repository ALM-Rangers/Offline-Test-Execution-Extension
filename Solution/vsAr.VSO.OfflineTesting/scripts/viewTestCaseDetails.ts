

import Controls = require("VSS/Controls");
import Combos = require("VSS/Controls/Combos");
import Grids = require("VSS/Controls/Grids");
import Menus = require("VSS/Controls/Menus");
import Dialogs = require("VSS/Controls/Dialogs");



import Common = require("scripts/Common");
import SvcExcel = require("scripts/serviceExcel");

export class viewTestCaseDetails{
    protected tesCase:SvcExcel.IImportData
    protected _content;
    public width: number = 700;
    public height: number = 400;;
    public Init(content, tc: SvcExcel.IImportData) {
        var view = this;
        view.tesCase = tc;
        view._content = content;

        if (tc.validWarnMsg) {
            view._content.find("#showWarningText").text(tc.validWarnMsg);
            view._content.find("#warningMsg").show();
        }
        if (tc.validErrMsg) {
            view._content.find("#showErrText").text(tc.validErrMsg);
            view._content.find("#errMsg").show();
        }


        view._content.find("#testCaseOutcome").text(tc.outcome);
        view._content.find("#testCaseConfig").text(tc.config);
        view._content.find("#testCaseTester").text(tc.tester);
        var $stepGrid = view._content.find("#gridTestSteps");
        var $rowTmple = view._content.find("#tmpltRowStep");
        var h = 250;
        if (tc.steps != null) {
            tc.steps.forEach(s => {
                var $row = $rowTmple.clone();
                var d = $("<div />");
                var dIcon = $("<div class='testpoint-outcome-shade icon bowtie-icon-small'/>");
                dIcon.addClass(Common.getIconFromTestOutcome(s.outcome));
                d.append(dIcon);
                var dTxt = $("<span />");
                dTxt.text(s.outcome);
                d.append(dTxt);

                $row.find("#index").text(s.index);
                $row.find("#action").text(s.action);
                $row.find("#expected").text(s.expectedResult);
                $row.find("#outcome").html(d.html());
                $row.find("#comment").text(s.comment);

                h += 22;
                $stepGrid.append($row);
            });

            $rowTmple.hide();
            this.height = h;
        }
        else {
            //$rowTmple.find("#gridTestSteps").html("<tr><td colspan="5">No test steps found</td></tr>");
        }
    }   
}


export function openTestCaseDetails( tc:SvcExcel.IImportData): IPromise<any> {
    var deferred = $.Deferred<any>();

    var view = this;
    var extensionContext = VSS.getExtensionContext();

    var $dlgContent = $("#showTestCaseDetailsView").clone();
    $dlgContent.show();
    $dlgContent.find("#showTestCaseDetailsView").show();


    var viewDlg: viewTestCaseDetails = new viewTestCaseDetails();

    viewDlg.Init($dlgContent, tc);

    var dlgOptions: Dialogs.IModalDialogOptions = {
        width: viewDlg.width,
        height: viewDlg.height,
        title: "Test case "+tc.testCaseId + ": "+ tc.title,
        content: $dlgContent,
      //  buttons: [{ text: "Close" }],
        
        okCallback: (result: any) => {
          
            deferred.resolve(null);
        }
    };

    var dialog = Dialogs.show(Dialogs.ModalDialog, dlgOptions);
    dialog.updateOkButton(true);
    dialog.setDialogResult(true);

    return deferred.promise();
}