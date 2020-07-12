/// <reference types="vss-web-extension-sdk" />

//---------------------------------------------------------------------
// <copyright file="serviceTestCase.js">
//    This code is licensed under the MIT License.
//    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF 
//    ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED 
//    TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A 
//    PARTICULAR PURPOSE AND NONINFRINGEMENT.
// </copyright>
// <summary>A service so get the resutl of work item queries as xml .</summary>
//---------------------------------------------------------------------

import StatusIndicator = require("VSS/Controls/StatusIndicator");
import Navigation = require("VSS/Controls/Navigation");
import UtilsHTML = require("VSS/Utils/Html");

import WebApi_Constants = require("VSS/WebApi/Constants");

import WorkItemContracts = require("TFS/WorkItemTracking/Contracts");
import WorkItemClient = require("TFS/WorkItemTracking/RestClient");
import WorkItemServices = require("TFS/WorkItemTracking/Services");

import CoreUtils = require("VSS/Utils/Core");
import UtilsDate = require("VSS/Utils/Date");
import UtilsString = require("VSS/Utils/String");

import TestClient = require("TFS/TestManagement/RestClient");
import TestContracts = require("TFS/TestManagement/Contracts");

import OperationClient = require("VSS/Operations/RestClient");
import OperationContract = require("VSS/Operations/Contracts");

import Q = require("q");

import Common = require("scripts/Common");
import QueryFetcher = require("scripts/QueryFetcher");
import TestHelper = require("scripts/TestHelper");



export interface IOffLineTestResult {
    testCaseId: number,
    testPointId: number,
    config: string,
    outcome?: string,
    comment?: string,
    execDate?: Date,
    iteration?: number,
    steps?: TestHelper.ITestStep[],
    parameters?: { [key: string]: string },
    runBy?: any
}

export function getConfigsForPlan():IPromise<TestContracts.TestConfiguration[]> {
    var tc = TestClient.getClient();
    
    return tc.getTestConfigurations(VSS.getWebContext().project.id);
}

export function getTestPlans(): IPromise<TestContracts.TestPlan[]> {
    var tc = TestClient.getClient();

    return tc.getPlans(VSS.getWebContext().project.name);
}

export function getTestPlan(id:number): IPromise<TestContracts.TestPlan> {
    var tc = TestClient.getClient();

    return tc.getPlanById(VSS.getWebContext().project.name, id);
}



export function getTestPointsForSuite(planId, startSuiteId, includeChilds: boolean):IPromise<TestContracts.TestPoint[]> {
    var deferred = $.Deferred<TestContracts.TestPoint[]>();
    var self = this;
    var tstClnt = TestClient.getClient();
    
    tstClnt.getTestSuitesForPlan(VSS.getWebContext().project.id, planId, includeChilds).then(testSuites => {
        var rootTestSuite: TestContracts.TestSuite = null;
        var suiteList: TestContracts.TestSuite[] = null;
        if (startSuiteId == null) {
            rootTestSuite = testSuites[0];
            suiteList = testSuites;
        }
        else {
            rootTestSuite = testSuites.filter(i=> { return i.id === startSuiteId })[0];
            suiteList = FindChildSuites(rootTestSuite, testSuites, [rootTestSuite])
        }
        if (!includeChilds) {
            suiteList = [rootTestSuite];
        }
        var prms: IPromise<TestContracts.TestPoint[]>[] = [];

        suiteList.forEach(suite => {
            
            prms.push(tstClnt.getPoints(suite.project.name, planId, suite.id, "System.Title"));
        });

        var lst: TestContracts.TestPoint[] = [];

        Q.all(prms).then(data => {
            data.forEach(tps => {
                lst= lst.concat(tps);
            });
            deferred.resolve(lst);
        });
                    
            
            
    });
    return deferred.promise();
}


    function FindChildSuites(suite, inputlist, list) {
        var childs = inputlist.filter(i=> { return i.parent != null && i.parent.id === suite.id.toString() });
        list = list.concat(childs);
        childs.forEach(function (c) {
            list = FindChildSuites(c, inputlist, list);
        })
        return list;
    }


    export class TestResultImporter {
        protected _tcClient: TestClient.TestHttpClient3_2 = null;
        protected msgLst: string[] = [];
        protected interval: number;

        constructor() {
            this._tcClient = TestClient.getClient(TestClient.TestHttpClient3_2 );
         }

          protected testPointOutcome(iterations: IOffLineTestResult[]): string {
            //Calculates the test point outcome from the iterations outcomes...
            var calcOutcome = new TestHelper.TestOucomeAggregator();

            iterations.forEach(i => {
                calcOutcome.addOutcome(i.outcome);
            });

            return calcOutcome.getOutcome();
        }

          public CreateTestRun(planId: number, runName: string, startDate: Date, idList: number[]): IPromise<TestContracts.TestRun> {

              var runModel: TestContracts.RunCreateModel = <TestContracts.RunCreateModel><any>{
                  automated: false,
                  //pointIds: idList,
                  name: runName,
                  plan: { "id": planId.toString(), name: "", url: null },
                  startDate: startDate.toISOString(),
                  state: "InProgress",
              }

              return this._tcClient.createTestRun(runModel, VSS.getWebContext().project.name);
          }

          public CompleteTestRun(runId: number, startDate: Date, endDate: Date): IPromise<TestContracts.TestRun> {

              var runModel: TestContracts.RunUpdateModel = <TestContracts.RunUpdateModel><any>{
                  deleteInProgressResults: true,
                  startedDate: startDate.toISOString(),

                  completedDate: endDate.toISOString(),
                  state: "Completed",
              };

              return this._tcClient.updateTestRun(runModel, VSS.getWebContext().project.name, runId);
          }

          public deleteTestRun(runId: number) {
              return this._tcClient.deleteTestRun(VSS.getWebContext().project.name, runId)
          }

          public publishTestResults(projName: string, runId: number, startDate: Date, endDate: Date, resulLst: IOffLineTestResult[]): IPromise<any> {
              var deferred = $.Deferred<any>();
              var self = this;

            var results = this.creatTestResulModel(startDate, endDate, resulLst);
            console.log("publishTestResults - model to publish");
            console.log(results);

            var resultsJson = JSON.stringify(results);
            
            var resultDoc: TestContracts.TestResultDocument = <TestContracts.TestResultDocument><any>{
                payload: {
                    name: "ResultDoc",
                    comment: "Imported Offline test execution results",
                    stream: btoa(resultsJson) // encodeURIComponent(resultsJson))
                }
            };

              self._tcClient.publishTestResultDocument(resultDoc, projName, runId).then(
                  data => {
                      console.log("Sent resultDocument to server, waiting for async serverjob to complete")
                      self.WaitForOperationToEnd(data.operationReference.id).then(
                          data => { deferred.resolve(data); },
                          err => { deferred.reject(err); }
                      );
                  },
                  err => {
                      console.log("Failed publishTestResult Document");

                      deferred.reject(err);
                  }
              );

              return deferred.promise();
          }

        public creatTestResulModel( startDate: Date, endDate: Date, resulLst: IOffLineTestResult[]): TestContracts.TestCaseResult[]  {
            var self = this;
          
            var trList: TestContracts.TestCaseResult[] = [];

            var testPointIdLst = resulLst.map(i => { return i.testPointId; });
            testPointIdLst = testPointIdLst["unique"]();

            testPointIdLst.forEach(tpId => {

                var oteIterations = resulLst.filter(i => { return i.testPointId === tpId; });
                var tpIteration = oteIterations[0];
                var trUpdate: TestContracts.TestCaseResult = <TestContracts.TestCaseResult><any>{

                    testPoint: {id: tpIteration.testPointId.toString()},
                    comment: tpIteration.comment,
                    completedDate: tpIteration.execDate != null ? tpIteration.execDate.toISOString() : endDate.toISOString(),

                    outcome: tpIteration.iteration? self.testPointOutcome(oteIterations): tpIteration.outcome,
                    startedDate: tpIteration.execDate != null ? tpIteration.execDate.toISOString() : startDate.toISOString(),
                    iterationDetails:[]
                };

                oteIterations.forEach(it => {
                    var iteration = self.CreateIterationModel(it);
                    // OK NOW Create Test Iterations
                    trUpdate.iterationDetails.push(iteration)
                });
                trList.push(trUpdate);

            });
            console.log("Built testResultModel:");
            console.log(JSON.stringify(trList));
            return trList;
        }

        protected WaitForOperationToEnd(operationId:string):IPromise<any> {
            var deferred = $.Deferred<any>();
            var self = this;

            self.interval = setInterval(
                () => {
                    console.log("Checking operation status");
                    var opClient = OperationClient.getClient();
                    opClient.getOperation(operationId).then(
                        status => {
                            if (status.status === OperationContract.OperationStatus.Succeeded) {
                                console.log("Operation status Succeeded ");
                                clearInterval(self.interval);
                                deferred.resolve(true);
                            }
                            else if (status.status === OperationContract.OperationStatus.Failed) {
                                console.log("Operation status Failed ");
                                clearInterval(self.interval);
                                deferred.reject(status);
                            }
                            else if (status.status === OperationContract.OperationStatus.Cancelled) {
                                console.log("Operation status Cancelled ");
                                clearInterval(self.interval);
                                deferred.reject(status);
                            }
                        }
                    );
                },
                3000);

            return deferred.promise();
        }

        protected CreateIterationModel(otr: IOffLineTestResult): TestContracts.TestIterationDetailsModel {

            var steps: TestContracts.TestActionResultModel[] = [];
            var params: TestContracts.TestResultParameterModel[] = [];
            if (otr.steps) {
                otr.steps.forEach(s => {
                    var stepIdentifier:string = s.id.toString();
                    if (s.parentStepId != null) {
                        stepIdentifier = s.parentStepId + ";" + s.id;
                    }

                    var tarm = <TestContracts.TestActionResultModel><any>{
                        outcome: s.outcome,
                        comment: s.comment,
                        errorMessage: s.comment,
                        stepIdentifier: stepIdentifier,
                    };
                    if (s.sharedStepWorkItemId){
                        tarm.sharedStepModel = { id: s.sharedStepWorkItemId, revision: s.sharedStepWorkItemRevision}
                    }

                    if (s.params) {
                        for (var p in s.params) {
                            if (s.params.hasOwnProperty(p)) {
                                params.push(<TestContracts.TestResultParameterModel><any>{
                                    parameterName: p,
                                    value: otr.parameters[p],
                                    stepIdentifier: stepIdentifier
                                })
                            }
                        }
                    }

                    steps.push(tarm)
                });
            }

            var iteration: TestContracts.TestIterationDetailsModel = <TestContracts.TestIterationDetailsModel ><any>{
                actionResults: steps,
                comment: otr.comment,
                outcome: otr.outcome,
            }
            if (params.length > 0) {
                iteration.parameters = params;
            }

            return iteration;
        }

        public GetTestPlan(projName: string, planId: number): IPromise<TestContracts.TestPlan> {
            return this._tcClient.getPlanById(projName, planId);
        }

        public fetchTestPoints(planId: number, lst: number[]): IPromise<TestContracts.TestPoint[]> {
            return this._tcClient.getPoints(VSS.getWebContext().project.id, planId, null, null, null, null, lst.join(","), true);
        }

        public fetchTestCases(planId: number, lst: number[], testHelper: TestHelper.TestHelper): IPromise<any[]> {
            var deferred = $.Deferred<any[]>();
            var self = this;
            var bulkFetcher = new QueryFetcher.BulkWIFetcher();
            bulkFetcher.FetchAllWorkItems(lst, ["System.Id", "System.Title", "Microsoft.VSTS.TCM.Steps", "Microsoft.VSTS.TCM.LocalDataSource"], 0).then(
                tcData => {
                    var sharedStepsIds = TestHelper.scanSharedStepsAndParameters(tcData, true, true);

                    var sharedBulkFetcher = new QueryFetcher.BulkWIFetcher();
                    sharedBulkFetcher.FetchAllWorkItems(sharedStepsIds, null, WorkItemContracts.WorkItemExpand.Fields).then(
                        sharedData => {
                            testHelper.setSharedWIData(sharedData);
                            deferred.resolve(tcData);
                        },
                        err => {
                            console.log("Error Prefetcing Shared items");
                            deferred.reject(err);
                        }
                    );
                },
                err => {
                    deferred.reject(err);
                });

            return deferred.promise();
        }     
    }
