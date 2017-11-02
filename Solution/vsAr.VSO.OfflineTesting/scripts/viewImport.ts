/// <reference types="vss-web-extension-sdk" />

import Dialog = require("VSS/Controls/Dialogs");
import Grids = require("VSS/Controls/Grids");
import FileInput = require("VSS/Controls/FileInput");
import Controls = require("VSS/Controls");
import Combos = require("VSS/Controls/Combos");
import Menus = require("VSS/Controls/Menus");

import TestClient = require("TFS/TestManagement/RestClient");
import VSS_Service = require("VSS/Service");
import TestContracts = require("TFS/TestManagement/Contracts");


import Q = require("q");


import Common = require("scripts/Common");
import SvcExcel = require("scripts/serviceExcel");
import SvcTest = require("scripts/serviceTestCase");
import QueryFetcher = require("scripts/QueryFetcher");
import TestHelper = require("scripts/TestHelper");
import TeleMetry = require("scripts/TelemetryClient");
import TestCaseDetailsDlg = require("scripts/viewTestCaseDetails");

export class viewImport extends Common.viewBase {
    public testResults: SvcExcel.IImportData[]=[];
    
    protected importedTestPlan: string;
    protected testPlan: TestContracts.TestPlan;
    protected testSuite: TestContracts.TestSuite;

    protected startDateCtrl: Combos.Combo;
    protected endDateCtrl: Combos.Combo;
    private context: WebContext;
    protected currentItem: SvcExcel.IImportData;
    protected grid: Grids.Grid;
    protected fileInput: FileInput.FileInputControl;

    protected onGoingValidations: string;
    protected orgValidSummary: JQuery;


    public updateContext(ctx) {
        var self = this;
        self.testPlan = ctx.selectedPlan;
        self.testSuite = ctx.selectedSuite;
    }


    public init() {
        var self = this;
        
        var fileInput: FileInput.FileInputControl = null;
        var startDateCtl: any = null;
        var endDateCtl: any = null;

        self.grid= self.initGrid();

        self.orgValidSummary = $("#validationSummary").clone();

        var options: FileInput.FileInputControlOptions = {
            maximumNumberOfFiles: 1,
            maximumSingleFileSize: 1e6,
            detectEncoding: true,
            allowedFileExtensions: ["xlsx"], 
            updateHandler: e => {
                self.uploadHandler(e)

            }
        };

        $("#cmdImport").click(e => {
            self.DoImport();
        });
        $("#cmdImport").prop("disabled", true);
        
        self.fileInput = Controls.create(FileInput.FileInputControl, $("#uploadCases"), options);

        var opts: Combos.IComboOptions = {
            type: "date-time"
        }
        self.startDateCtrl = Controls.create(Combos.Combo, $("#dtStartDate"), opts);
        self.endDateCtrl = Controls.create(Combos.Combo, $("#dtEndDate"), opts);

        //TestCode
        

}

    protected initGrid() {
        var self = this;

        var gutterOpts: Grids.IGridGutterOptions = {
            contextMenu: true
        }

         //"TestCaseId", "Title", "TestPointId", "Configuration", "Tester", "TestStepId", "StepAction", "StepExpected",  "OutCome", "Comments"]
        var cols: Array<Grids.IGridColumn> = [
            { index: "validErrMsg", text: "Valid", getCellContents: getValidCellContent, width:40},
            { index: "testCaseId", text: "Test Case ID" },
            { index: "title", text: "Title" },
            { index: "testPointId", text: "Test Point", getCellContents: getTestPointContent},
            { index: "", text: "Parameters" , getCellContents:getItrationContent},
            { index: "config", text: "Configuration" },
            { index: "tester", text: "Tester" },
            //{ index: "testStepId", text: "Test Step Id" },
            //{ index: "action", text: "Step Action" },
            //{ index: "expected", text: "Step Expected" },
            { index: "outcome", text: "Outcome", getCellContents: getOutcomeContent},
            { index: "comment", text: "Comment", width:200 },
            //{ index: "execDate", text: "Date" }
        ]

        var mnuItems: Array<Menus.IMenuItemSpec> = [
            { icon: "bowtie-status-success", text: "Passed", id: "pass" },
            { icon: "bowtie-status-failure", text: "Failed", id: "fail" },
            {
                icon: "bowtie-status-blocked", text: "Blocked", id: "block"
            }
        ]
        var gridOpts: Grids.IGridOptions = {
           
            width: "100%",
            height: "100%",
            contextMenu: {
                items: mnuItems,
                executeAction: self.menuItemClick
            },
            gutter: gutterOpts,
            lastCellFillsRemainingContent: false,
            columns: cols,
            openRowDetail: (index: number) => {
                
                self.currentItem = <SvcExcel.IImportData>self.grid.getRowData(index);
                
                TestCaseDetailsDlg.openTestCaseDetails(self.currentItem);
                
            }
        }

        return  Controls.create(Grids.Grid, $('#importGrid'), gridOpts);
    }
    
    public uploadHandler(e: FileInput.FileInputControlUpdateEventData) {
        var self = this;
        $("#cmdImport").prop("disabled", true);

        self.grid.setDataSource(null);
        if (e.files.length == 1) {
            var impData = SvcExcel.importFromExcel(self.fileInput);
            
            self.importedTestPlan = impData.plan.replace(";", ":");
            $("#uploadHeader").hide();
            $("#uploadedHeader").show();
            $("#validationSummary").html(self.orgValidSummary.html());
            $("#validationSummary").show();
            $("#validatePane").show();
            $("#testRunResult").hide();

            self.testResults = impData.data;
            self.grid.setDataSource(self.testResults);
            $('#testPlan').text(self.importedTestPlan);
            $('#testPlan').css("color", "black");
            $('#txtRunName').val("Offline test run " + new Date().toLocaleString());

            self.startDateCtrl.setText(impData.minDate != null ? impData.minDate.toDateString() : new Date().toDateString());
            self.endDateCtrl.setText(impData.maxDate != null ? impData.maxDate.toDateString() : new Date().toDateString());

            self.Validate(impData.containsSteps );
     } else {
            $("#uploadHeader").show();
            $("#uploadedHeader").hide();
            $("#validationSummary").hide();
            $("#validatePane").hide();

            $("#content-container").empty();
            $('#navToolbarImport').hide();
            $("#testRunResult").hide();
        }

    }


    public menuItemClick(args) {

        switch (args.get_commandName()) {
            case "block":
                alert(JSON.stringify(this.currentItem));
                break;
            case "visualize":
                break;
        }
    }

    public DoImport() {
        var self = this;
        self.StartLoading(true, "Beginning import...");
        var planId = Number(self.importedTestPlan.split(":")[0]);
        var startDate = new Date(self.startDateCtrl.getText());
        var endDate = new Date(self.endDateCtrl.getText());
        var runName = $("#txtRunName").val();
        self.ImportTestPoints(planId, runName, startDate, endDate).then(
            data => {
                TeleMetry.TelemetryClient.getClient().trackEvent("Import Success", { testPointCnt: self.testResults.length });
                self.DoneLoading();
                $("#success").show();
                $("#createTestRunParam").hide();
                $("#validatePane").hide();
                $("#testRunResult").show();
                var msg = $("#resultMsg").text();
                msg = msg.replace("{testCount}", data.totalTests.toString());
                $("#resultMsg").text(msg);
                $("#testRunLnk").text("# " + data.id);
                $("#testRunLnk").attr("href", data.webAccessUrl);
            },
            err => {
                TeleMetry.TelemetryClient.getClient().trackEvent("Import Error", { error: err, testPointCnt: self.testResults.length });
                self.DoneLoading();
                $("#failure").show();
                $("#createTestRunParam").hide();
                $("#validatePane").hide();
                $("#testRunResult").show();
                $("#resultMsg").text(err);
                $("#testRunDiv").hide();
                
            }
        );
    }

    public ImportTestPoints(planId: number, testRunName: string, startDate: Date, endDate: Date):IPromise<TestContracts.TestRun> {
        var deferred = $.Deferred<TestContracts.TestRun>();
        var self = this;

        var proj = VSS.getWebContext().project.id;

        var importer = new SvcTest.TestResultImporter()

        self.ProgressUpdate("Creating testrun");
        console.log("Creating testrun");
        importer.CreateTestRun(planId, testRunName, startDate, self.testResults.map(i => { return i.testPointId; })).then(
            testRun => {
                console.log("Testrun created id= " + testRun.id);
                self.ProgressUpdate("Updating test points");
                importer.publishTestResults(proj, testRun.id, startDate, endDate, self.testResults).then(
                    testResults => {
                        console.log("Test points update");
                        self.ProgressUpdate("Closing test run");
                        importer.CompleteTestRun(testRun.id, startDate, endDate).then(
                            testRun => {
                                console.log("Test run closed - all done");
                                deferred.resolve(testRun);
                            },
                            err => {
                                console.log("Test run close - failed");
                                console.log(err);
                                deferred.reject(err);
                            }
                        );
                    },
                    errPublish => {
                        importer.deleteTestRun(testRun.id);
                        console.log("publishTestResults error");
                        console.log(errPublish);
                        deferred.reject(errPublish);
                    }
                );
            },
            err => {
                console.log("CreateTestRun error" );
                console.log(err);
                deferred.reject(err);

            }
        );

        return deferred.promise();
    }

    public Validate(containsSteps:boolean) {
        var self = this;
        
        self.onGoingValidations = "Validating Outcome, Test cases, Test points";
        self.StartLoading(true,  self.onGoingValidations);

        var prms: IPromise<any>[] = [
            self.ValidatePlan(),
            self.ValidateTestPoints(),
            self.ValidateOutcomeAndDuplicates(),
            self.ValidateTestCases(containsSteps)];
        Q.all(prms).then(data => {
            var prop = {
                testPointCnt: self.testResults.length,
                validPlan: data[0],
                validTestPoints: data[1],
                validOutcome: data[2],
                validTestCases: data[3],

            }

            TeleMetry.TelemetryClient.getClient().trackEvent("Validation", prop);
            self.DoneLoading();
            self.grid.setDataSource(self.testResults);

            //if all validation pass without errors... 
            if (data[0] && data[1] && data[2] && data[3]) {
                $("#cmdImport").prop("disabled", false);
            }
        });
    }
    
    protected setValidmsg(status: "passed" | "warning" | "error", msg: string, imgSelector: string, txtSelector : string ) {
        $(imgSelector).attr("src", "img/icon-"+ status+".png");
        $(txtSelector).text(msg);
    }

    public ValidatePlan(): IPromise<any> {
        var deferred = $.Deferred<any>();

        var self = this;
        var importedTestPlanId:number = Number(self.importedTestPlan.split(":")[0]);
        if (importedTestPlanId== self.testPlan.id) {
            self.importedTestPlan = self.testPlan.id + ":" + self.testPlan.name;
            self.setValidmsg("passed", "Imported data is from this test plan", "#validTestPlanImg", "#validTestPlan");
            deferred.resolve(true);
        }
        else {
            $('#testPlan').text(self.importedTestPlan);
            $('#testPlan').css("color", "red");
            
            SvcTest.getTestPlan(importedTestPlanId).then(
                tp => {
                    self.setValidmsg("warning", "Warning - the imported data is not for the active test plan","#validTestPlanImg", "#validTestPlan" );
                    deferred.resolve(true);
                },
                err => {
                    self.setValidmsg("error", "Can /'t find test plan id " + importedTestPlanId, "#validTestPlanImg", "#validTestPlan");
                    deferred.resolve(false);
                });
        }
        return deferred.promise();
    }

    public ValidateTestPoints(): IPromise<any> {
        var deferred = $.Deferred<any>();
        var self = this;
        var noTestPointMissing = 0;
        var warningCnt = 0;
        var missingTestPoints: number[] = [];
        SvcTest.getTestPointsForSuite(self.testPlan.id, self.testSuite.id, true).then(
            tpLst => {
                self.testResults.forEach(r => {
                    var foundTP = tpLst.filter(i => { return i.id == r.testPointId })[0];
                    if (foundTP != null) {
                        var foundTester = foundTP.assignedTo != null ? Common.getDisplayName(foundTP.assignedTo) : ""
                        var foundConfig = foundTP.configuration != null ? foundTP.configuration.name : "";

                        if (foundConfig != r.config) {
                            r.validWarnMsg = "Config doesnt match (" + r.config + "!=" + foundConfig + ")";
                            warningCnt++;
                        }
                        if (foundTester!= r.tester) {
                            r.validWarnMsg += r.validWarnMsg!=""?" ":"" +  "Tester doesnt match (" + r.tester +"!="+ foundTester+")";
                            warningCnt++;
                        }
                    }
                    else {
                        r.validWarnMsg = "Test point not found in current context";
                        missingTestPoints.push(r.testPointId);
                        noTestPointMissing++;
                        warningCnt++;
                    }
                });


                if (self.testResults.length == 0) {
                    self.setValidmsg("error", "No valid test point to import found", "#validTestPointsImg", "#validTestPoints");
                }
                else if (noTestPointMissing==0 && warningCnt == 0) {
                    self.setValidmsg("passed", "All test points found and valid within selected context", "#validTestPointsImg", "#validTestPoints");
                }
                else {
                    var msg = "";
                    if (noTestPointMissing > 0) {
                        msg = noTestPointMissing + " test points not found in current context";
                    }
                    if (warningCnt > 0) {
                        msg += msg!=""?". ":"" + warningCnt + " test points with warnings";
                        
                    }
                    self.setValidmsg("warning", msg, "#validTestPointsImg", "#validTestPoints");
                }

                self.onGoingValidations = self.onGoingValidations.replace(", Test Point", "");
                self.ProgressUpdate(self.onGoingValidations);
                deferred.resolve(true);
            },
            err => {
                deferred.resolve(false);
            }
        );
      
        return deferred.promise();
    }

    public ValidateTestCases(validateTestStep: boolean):IPromise<any> {
        var deferred = $.Deferred<any>();
        var self = this;
        var tpFetcher = new SvcTest.TestResultImporter();
        var planId = Number(this.importedTestPlan.split(":")[0]);
        var warningCnt: number = 0;
        var errorCnt: number = 0;
        var testHelper: TestHelper.TestHelper = new TestHelper.TestHelper();

        tpFetcher.fetchTestCases(planId, this.testResults.map(i => { return i.testCaseId; }), testHelper).then(
                data => {
                self.testResults.forEach(i => {
                    var r = self.ValidateTC(validateTestStep, i, data, testHelper);
                    warningCnt += r.warning;
                    errorCnt += r.errors;

                });
                self.onGoingValidations = self.onGoingValidations.replace(", Test cases", "");
                self.ProgressUpdate(self.onGoingValidations);

                if (warningCnt == 0 && errorCnt == 0) {
                    self.setValidmsg("passed", "All test case data is valid", "#validTestCasesImg", "#validTestCases");
                }
                else if (errorCnt > 0) {
                    self.setValidmsg("error", "Found " + errorCnt + " errors in test case data", "#validTestCasesImg", "#validTestCases");
                }
                else if (warningCnt > 0) {
                    self.setValidmsg("warning", "Found " + warningCnt + " warnings in test case data", "#validTestCasesImg", "#validTestCases");
                }
                deferred.resolve(errorCnt==0);
            }
        );

        return deferred.promise();
    }

    public ValidateOutcomeAndDuplicates():IPromise<any> {
        var deferred = $.Deferred<any>();
        var self = this;
        var errorCnt:number = 0;
        var okOutComes = "Passed|Failed|Blocked|Paused|";
        self.testResults.forEach(i => {
            if (okOutComes.indexOf(i.outcome) == -1) {
                i.validErrMsg = "Not supported outcome";
                i.validErrMsg = "Not supported outcome";
                errorCnt++;
            }
            if (self.testResults.filter(j => { return j.testPointId === i.testPointId && j.iteration===i.iteration; }).length > 1) {
                i.validErrMsg += "Duplicated test point id and iterations found";
                errorCnt++;
            }
        });
        self.onGoingValidations = self.onGoingValidations.replace("Outcome,", "");
        self.ProgressUpdate(self.onGoingValidations);

        if (errorCnt == 0) {
            self.setValidmsg("passed", "All test outcome data is valid", "#validOutcomeImg", "#validOutcome");
        }
        else  {
            self.setValidmsg("error", "Found " + errorCnt + " errors in test outcome data", "#validOutcomeImg", "#validOutcome");
        }
        console.log("ValidateOutcome done");
        deferred.resolve(errorCnt==0);

        return deferred.promise();
    }
    
    public ValidateTC(validateTestStep:boolean,  importRow: SvcExcel.IImportData, tcdata, testHelper:TestHelper.TestHelper): { warning: number, errors: number }{
        var warningCnt: number = 0;
        var errorsCnt: number = 0;
     

        var tc = tcdata.filter(i => { return i.id == importRow.testCaseId; })[0];
        if (tc != null) {
            if (importRow.title !=tc.fields['System.Title']){
                importRow.validWarnMsg = "Test case title doesnt match";
                warningCnt++;
            }
            if (importRow.steps != null) {
                var testStepsValidation = this.ValidateTestSteps(importRow, tc,testHelper);
                errorsCnt += testStepsValidation.errors;
                warningCnt += testStepsValidation.warning;
            }
            else {
                if (validateTestStep && tc.fields['Microsoft.VSTS.TCM.Steps'] != null) {
                    importRow.validErrMsg = "Test case doesnt contain the test steps";
                    errorsCnt++;
                }
            }
        }
        else {
            importRow.validErrMsg = "Test case not found";
            errorsCnt++;
        }
        return { warning: warningCnt, errors: errorsCnt };
    }

    protected ValidateTestSteps(importRow: SvcExcel.IImportData, tc, testHelper: TestHelper.TestHelper ) {
        var warningCnt: number = 0;
        var errorsCnt: number = 0;

        var tcSteps = testHelper.parseTestCaseSteps(tc, importRow.parameters);
        if (importRow.steps.length != tcSteps.length) {
            importRow.validErrMsg = "Test case doesnt contain the same number of test steps";
            errorsCnt++;
        }
        var calcOutcome = new TestHelper.TestOucomeAggregator();
        importRow.steps.forEach(importStep => {
            var orgStep = tcSteps.filter(i => { return i.index == importStep.index })[0];
            if (orgStep != null) {
                importStep.id = orgStep.id;
                importStep.parentStepId = orgStep.parentStepId;
                importStep.sharedStepWorkItemId = orgStep.sharedStepWorkItemId;
                importStep.sharedStepWorkItemRevision = orgStep.sharedStepWorkItemRevision;
                importStep.params = orgStep.params;
                if (importStep.outcome != "") {
                    calcOutcome.addOutcome(importStep.outcome);
                }
            }
            else {
                importRow.validWarnMsg = "Test step not found (index" + importStep.index + ")";
                warningCnt++;
            }
        });
        if (calcOutcome.getOutcomeCount() > 0) {
            if (importRow.outcome === "") {
                importRow.outcome = calcOutcome.getOutcome();
            }
            else {
                if (importRow.outcome != calcOutcome.getOutcome()) {
                    importRow.validWarnMsg = "Test case outcome differes from the aggregated test steps outcome";
                    warningCnt++;
                }
            }
        }
        return { warning: warningCnt, errors: errorsCnt };
    }
}


function getValidCellContent(rowInfo, dataIndex, expandedState, level, column, indentIndex, columnOrder) {
    var err = this.getColumnValue(dataIndex, "validErrMsg");
    var warn = this.getColumnValue(dataIndex, "validWarnMsg");
    var $d = $("<div class='grid-cell'/>").width(column.width || 100)

    var title = "";
    if (err != null) {
        var dIcon = $("<div class='icon bowtie-icon-small'/>");
        dIcon.addClass("icon-valid-error");
        $d.append(dIcon);
        title = title + " " + err ;
    }

    if (warn != null) {
        var dIcon = $("<div class='icon bowtie-icon-small'/>");
        dIcon.addClass("icon-valid-warning");
        $d.append(dIcon);
        title = title +" " +  warn;
    }
      
    if (err == null && warn==null) {
        var dIcon = $("<div class='icon bowtie-icon-small'/>");
        dIcon.addClass("icon-valid-passed");
        $d.append(dIcon);
        
    }
    $d.prop('title', title);
    return $d;
}

function getOutcomeContent(rowInfo, dataIndex, expandedState, level, column, indentIndex, columnOrder) {
    
    var outcome = this.getColumnValue(dataIndex, column.index);
    var d = $("<div class='grid-cell'/>").width(column.width || 100)
    var dIcon = $("<div class='testpoint-outcome-shade icon bowtie-icon-small'/>");
    dIcon.addClass(Common.getIconFromTestOutcome(outcome));
    d.append(dIcon);
    var dTxt = $("<span />");
    dTxt.text(outcome);
    d.append(dTxt);
    return d;
}

function getItrationContent(rowInfo, dataIndex, expandedState, level, column, indentIndex, columnOrder) {

    var params: { [key: string]: string } = <{ [key: string]: string }>this._dataSource[dataIndex].parameters;

    var s = "";
    for (var p in params) {
        if (params.hasOwnProperty(p)) {
            s += "@" + p;
            s += "=" + params[p];
            s += "; ";
        }
    }

    var $d = $("<div class='grid-cell'/>").width(column.width || 100)
    $d.text(s)
    return $d;
}

function getTestPointContent(rowInfo, dataIndex, expandedState, level, column, indentIndex, columnOrder) {

    var tp = this.getColumnValue(dataIndex, column.index);
    var it = this.getColumnValue(dataIndex, "iteration");
    if (it != null) {
        tp += ":" + it;
    }

    var $d = $("<div class='grid-cell'/>").width(column.width || 100)
    $d.text(tp );
    return $d;
}
