/// <reference types="vss-web-extension-sdk" />


import Dialog = require("VSS/Controls/Dialogs");
import Grids = require("VSS/Controls/Grids");
import FileInput = require("VSS/Controls/FileInput");
import Controls = require("VSS/Controls");
import Combos = require("VSS/Controls/Combos");
import Menus = require("VSS/Controls/Menus");
import APIContracts = require("VSS/WebApi/Contracts");
import VSS_Service = require("VSS/Service");
import TestContracts = require("TFS/TestManagement/Contracts");
import Q = require("q");

import Common = require("scripts/Common");
import SvcExcel = require("scripts/serviceExcel");
import SvcTest = require("scripts/serviceTestCase");
import QueryFetcher = require("scripts/QueryFetcher");
import TestHelper = require("scripts/TestHelper");
import TeleMetry = require("scripts/TelemetryClient");


var str_ALL: string = "-- All --";
var str_PLAN: string = "Entire Test Plan";
var str_SuiteOnly: string = "Selected suite only";
var str_SuiteAndChilds: string = "Selected suite and child suites";


export class viewExport extends Common.viewBase {
    public testResults: SvcExcel.IImportData[] = [];

    protected testPlan: TestContracts.TestPlan;
    protected testSuite: TestContracts.TestSuite;
    protected rootSuiteId: number;
    protected ctxIncludeChilds: boolean=true;

    protected cboInclude: Combos.Combo;
    protected cboTester: Combos.Combo;
    protected cboConfig: Combos.Combo;
    private context: WebContext;
    protected currentItem: SvcExcel.IImportData;
    protected grid: Grids.Grid;

    protected lstTP: SvcExcel.IImportData[] = [];
    protected testersList: string[] = [];

    public init() {
        var self = this;

        var optInclude: Combos.IComboOptions = {
            type: "list",
            mode: "drop",
            allowEdit: false,
            source: [str_PLAN, str_SuiteAndChilds, str_SuiteOnly],
            change: () => { self.updateTestPointGrid(); }
        }
        self.cboInclude = Controls.create(Combos.Combo, $('#cboInclude'), optInclude);
        self.cboInclude.setSelectedIndex(1);

        var optConfig: Combos.IComboOptions = {
            type: "list",
            mode: "drop",
            allowEdit: false,
            source:[str_ALL],
            change: () => { self.grid.setDataSource(self.getFilteredTP()); }
        }
        self.cboConfig = Controls.create(Combos.Combo, $('#cboConfig'), optConfig);
        self.cboConfig.setSelectedIndex(0);

        var optTester: Combos.IComboOptions = {
            type: "list",
            mode: "drop",
            allowEdit: false,
            source: [str_ALL],
            change: () => { self.grid.setDataSource(self.getFilteredTP()); }
        }
        self.cboTester = Controls.create(Combos.Combo, $('#cboTester'), optTester);
        self.cboTester.setSelectedIndex(0);

        self.grid = self.initGrid();


        SvcTest.getConfigsForPlan().then(configs => {
            var src: string[] = configs.map(i => { return i.name; });
            src.push(str_ALL);
            self.cboConfig.setSource(src);
        });

        $("#cmdExport").click(e => {
            var exportTestSteps:boolean = $("#chkIncludeTestSteps")[0].checked;
            self.export(exportTestSteps);
        });
    }

    protected export(exportTestSteps:boolean) {
        var self = this;

        self.StartLoading(true, "Exporting to excel...");
        var tabName = self.testPlan.id + ";" + self.testPlan.name;
        tabName = tabName.replace(/[*?:\/\[\]]/g, '');
        
        var tpLst = self.getFilteredTP();
        var tcIdList = tpLst.map(i => { return i.testCaseId; });

        var tcFetcher = new SvcTest.TestResultImporter();
        var testHelper: TestHelper.TestHelper = new TestHelper.TestHelper();


        tcFetcher.fetchTestCases(self.testPlan.id, tcIdList, testHelper).then(         
            tcLst => {
                tcLst.forEach(tc => {
                    if (tc.fields["Microsoft.VSTS.TCM.Steps"] != null) {
                        
                        var iterations = testHelper.parseTestParams(tc.fields['Microsoft.VSTS.TCM.LocalDataSource'])

                        if (iterations != null && iterations.params != null &&  iterations.params.length > 0) {
                            var nFoundIndex = -1;
                            var pointIterations: { iteration: number, steps: TestHelper.ITestStep[] }[]=[]
                            iterations.params.forEach((p, ix) => {
                                var steps = testHelper.parseTestCaseSteps(tc, p)
                                pointIterations.push({ iteration: ix, steps: steps }) 
                            });


                            //tpLst.forEach((tp, ix) => {
                            for (var ix = tpLst.length-1; ix >= 0; ix--) {
                                var tp = tpLst[ix];
                           
                                if (tp.testCaseId == tc.id) {
                                   

                                    pointIterations.forEach((pi, pi_ix)=> {
                                        var i: SvcExcel.IImportData = {
                                            testCaseId: tp.testCaseId,
                                            title: tp.title,
                                            testPointId: tp.testPointId,
                                            iteration: pi.iteration+1,
                                            config: tp.config,
                                            tester: tp.tester,
                                            outcome: null,
                                            comment: "",
                                            execDate: tp.execDate,
                                            steps: exportTestSteps ? pi.steps : null,
                                            runBy : tp.runBy
                                        };
                                        tpLst.splice(ix+pi_ix, (pi_ix==0)?1:0, i);
                                        
                                    });
                                   

                                    tp.iteration = 0;
                                    if (exportTestSteps) {
                                        tp.steps = steps;
                                    };
                                }
                            };
                        }
                        else {
                            
                            var steps = testHelper.parseTestCaseSteps(tc);
                            tpLst.forEach(tp => {
                                if (tp.testCaseId == tc.id) {
                                    tp.iteration = 0;
                                    if (exportTestSteps) {
                                        tp.steps = steps;
                                    }
                                }
                            });

                        }
                    }
                });

                var xlWB = SvcExcel.exportToExcel(tabName, exportTestSteps, tpLst, { bookType: 'xlsx', bookSST: true, type: 'binary', name: self.testSuite.name + ".xlsx" });
                var config = self.cboConfig.getText();
                var tester = self.cboTester.getText();
                var fileName = "OTE_" + Common.secureForFileName(self.testSuite.name) + "_" + Common.secureForFileName(config) + "_" + Common.secureForFileName(tester) + ".xlsx";

                Common.SaveFile(xlWB, fileName, "application/vnd.ms-excel");
                TeleMetry.TelemetryClient.getClient().trackEvent("Export Success", { testPointCnt: tpLst.length, config: config, tester: tester });
                self.DoneLoading();
            },
            err => {
                TeleMetry.TelemetryClient.getClient().trackException(err);
                self.DoneLoading();
            }
        );
    }

    public updateContext(ctx) {
        console.log("updateContext");
        console.log(ctx);

        var self = this;
        if (ctx.selectedPlan != null) {

            var planId = ctx.selectedPlan.id;
            var suiteId = ctx.selectedSuite.id;
            self.testPlan = ctx.selectedPlan;
            self.testSuite = ctx.selectedSuite;
            self.rootSuiteId = Number(ctx.selectedPlan.rootSuiteId);
            
            self.updateTestPointGrid();
         
        }
        else{
            console.log("empty context");
        }
    }

    protected updateTestPointGrid() {
        var self = this;
        if (self.testPlan != null) {
            var suiteId = self.testSuite.id
            var includeChilds = true;

            switch (self.cboInclude.getValue()) {
                case str_PLAN:
                    suiteId = self.rootSuiteId;
                    includeChilds = true;
                    
                    break;
                case str_SuiteOnly:
                    includeChilds = false;
                    break;
                default:
                    includeChilds = true;
                    break;
            }


            self.LoadTestPoints(self.testPlan.id, suiteId, includeChilds).then(data => {
                console.log("Loaded testpoints " +suiteId +" includeChilds " + includeChilds);
                var lst = self.testersList;
                if (lst.indexOf(str_ALL) == -1) {
                    lst.push(str_ALL);
                }
                self.cboTester.setSource(lst);
                self.grid.setDataSource(self.getFilteredTP());

                TeleMetry.TelemetryClient.getClient().trackEvent("Grid refresh", { testPointCnt: self.getFilteredTP().length, include: self.cboInclude.getText(), tester:self.cboTester.getText(), config :self.cboConfig.getText() });

            });
        }
    }

    protected initGrid() {
        var self = this;

        var gutterOpts: Grids.IGridGutterOptions = {
            contextMenu: true
        }

        //"TestCaseId", "Title", "TestPointId", "Configuration", "Tester", "TestStepId", "StepAction", "StepExpected",  "OutCome", "Comments"]
        var cols: Array<Grids.IGridColumn> = [
            { index: "testCaseId", text: "Test Case", width: 75 },
            { index: "title", text: "Title", width:300 },
            { index: "testPointId", text: "Test Point", width: 75},
            { index: "config", text: "Configuration", width: 150},
            { index: "tester", text: "Tester", width: 150 },
            { index: "", text: "", width: 2 }
      
        ]

   
        var gridOpts: Grids.IGridOptions = {
            height: "100%",
            width: "100%",
    
            gutter: gutterOpts,
            lastCellFillsRemainingContent: false,
            columns: cols,
            openRowDetail: (index: number) => {

                self.currentItem = <SvcExcel.IImportData>self.grid.getRowData(index);
                var i = self.currentItem;
                var msg = i.testCaseId + " " + i.title + "\n";
                if (i.steps != null) {
                    i.steps.forEach(s => {
                        msg += "   " + s.action + " : " + s.expectedResult + " : " + s.outcome + "\n";
                    });
                }
                alert(msg);
            }
        }

        return Controls.create(Grids.Grid, $('#exportGrid'), gridOpts);
    }

    protected getFilteredTP() {
        var self = this;
        var fltConf = this.cboConfig.getText();
        if (fltConf == str_ALL) {
            fltConf = null;
        }
        var fltTst = this.cboTester.getText();
        if (fltTst == str_ALL) {
            fltTst = null;
        }

        return this.lstTP.filter(i => {
            return (fltConf == null ? true : i.config == fltConf) && (fltTst == null ? true : i.tester == fltTst);
        });
    }

    protected LoadTestPoints(planId:number, startSuiteId:number, includeChilds:boolean):IPromise<any> {
        var deferred = $.Deferred<any>();
        var self = this;
      //  var svc = new SvcTest.WorkItemQueryService();
        //svc.testPlanId = planId;

        //svc.teamProject = VSS.getWebContext().project.name;
        self.StartLoading(true, "Fetching testpoints ...");

        SvcTest.getTestPointsForSuite(planId, startSuiteId, includeChilds).then(lstTP => {

            self.lstTP = lstTP.map(i => {return convertToImportData(i);});
            self.lstTP.forEach(i => {
                if (self.testersList.indexOf(i.tester) == -1) {
                    self.testersList.push(i.tester);
                }
            });

            self.DoneLoading();
            deferred.resolve(true);
        });


        //svc.execute(this.ProgressUpdate.bind(self)).then((result: XMLDocument) => {
        //    var obj: any = Common.XML2jsobj(result.documentElement);
        //    self.lstTP = [];
        //    self.testersList = [];
        //    obj.testSuites.testSuite.forEach(suite => {
        //        if (suite.testCases.count > 1) {
        //            suite.testCases.testCase.forEach(tp => {
        //                var i = convertToImportData(tp)
        //                SvcExcel.addTestSteps(i, tp);
        //                self.lstTP.push(i);

        //                if (self.testersList.indexOf(i.tester) == -1) {
        //                    self.testersList.push(i.tester);
        //                }


        //            });
        //        } else if (suite.testCases.count == 1) {
        //            var tp = suite.testCases.testCase
        //            var i = convertToImportData(tp)
        //            SvcExcel.addTestSteps(i, tp);
        //            self.lstTP.push(i);

        //            if (self.testersList.indexOf(i.tester) == -1) {
        //                self.testersList.push(i.tester);
        //            }
        //        }

        //    });
        //    self.DoneLoading();
        //    deferred.resolve(true);
        //});

        return deferred.promise();
    }

}

function convertToImportData(tp:TestContracts.TestPoint): SvcExcel.IImportData {
    return {
        testCaseId: Number(tp.testCase.id),
        title: tp.workItemProperties[0].workItem.value,
        testPointId: tp.id,
        config: tp.configuration!=null? tp.configuration.name:"",
        tester: tp.assignedTo != null ? Common.getDisplayName(tp.assignedTo) : ""
     
    };
}


