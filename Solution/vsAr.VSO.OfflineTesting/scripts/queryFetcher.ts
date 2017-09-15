/// <reference types="vss-web-extension-sdk" />

//---------------------------------------------------------------------
// <copyright file="query-fetcher">
//    This code is licensed under the MIT License.
//    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF 
//    ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED 
//    TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A 
//    PARTICULAR PURPOSE AND NONINFRINGEMENT.
// </copyright>
// <summary>A utility to execute a query, fetch and build all work item data for the query result.</summary>
//---------------------------------------------------------------------

import Controls = require("VSS/Controls");
import VSS_Service = require("VSS/Service");
import WorkItemContracts = require("TFS/WorkItemTracking/Contracts");
import WorkItemClient = require("TFS/WorkItemTracking/RestClient");
import Utils_HTML = require("VSS/Utils/Html");
import CoreUtils = require("VSS/Utils/Core");
import WorkItemServices = require("TFS/WorkItemTracking/Services");

import Q= require("q");


import Common = require("scripts/Common");

export class QueryFetcher {

    private _progressCallback: Common.IProgressCallback;
    private _queryResult;
    private _queryData;
    private _queryResultList;
    private _treeRoot: any[];

    public executeQuery(queryId, progressCallback: Common.IProgressCallback):IPromise<any> {

        var deferred = $.Deferred<any>();

        var queryFetcher = this;
        this._progressCallback = progressCallback;

        var witClient = VSS_Service.getCollectionClient(WorkItemClient.WorkItemTrackingHttpClient);               
        queryFetcher.ProgressMessage("Executing query");
        console.log("Executing query");
        var context = VSS.getWebContext();

        var context = VSS.getWebContext();
        witClient.queryById(queryId, context.project.name, context.team.name).then(
            queryResult=> {
                console.log("Got query result");
                queryFetcher._queryResult = queryResult;

                var wi = [];
                var ids = [];
                if (queryResult.queryResultType == 1) {
                    ids = queryResult.workItems.map( wi => { return wi.id; });
                }
                if (queryResult.queryResultType == 2) {
                    ids = queryResult.workItemRelations.map(function (reference) { return reference.target.id; });
                }
                console.log("Fetch data");

                var bulkFetcher = new BulkWIFetcher()
                bulkFetcher.FetchAllWorkItems( ids, queryResult.columns.map(col => { return col.referenceName; }), WorkItemContracts.WorkItemExpand.Fields)
                    .then(
                    queryData => {
                            queryFetcher._queryData = queryData;
                            queryFetcher.ProgressMessage("Retrieved data, building hierarchy");
                            if (queryFetcher._queryResult.queryResultType == 2) {
                                queryFetcher._treeRoot = [];
                                var rootItems = queryFetcher._queryResult.workItemRelations.filter(function (o) { return o.source == undefined });

                                rootItems.forEach(function (n) {
                                    var workItem = queryFetcher._queryData.filter(function (wi) { return wi.id == n.target.id })[0];
                                    queryFetcher._treeRoot.push(workItem);
                                    try {
                                        queryFetcher.BuildTree(workItem);
                                    }
                                    catch (errBuildTree) {
                                        deferred.reject(errBuildTree);
                                    }

                                });
                                queryFetcher._queryResultList = queryFetcher._treeRoot;

                            }
                            if (queryFetcher._queryResult.queryResultType == 1) {
                                queryFetcher._queryResultList = queryFetcher._queryData;
                            }

                            deferred.resolve({ Columns: queryFetcher._queryResult.columns, QueryResults: queryFetcher._queryResultList });
                        },
                        err =>{
                            deferred.reject(err);
                        });
            },

            errGetById => {
                deferred.reject(errGetById);
            });
        
            return deferred.promise();
    }

  
        private BuildTree  ( wiParent) {
            this._queryResult.workItemRelations.filter( o=>  {
                if (o.source == null)
                    return false;
                return o.source.id == wiParent.id;
            })
                .forEach( item => {
                    if (item.childs == null) {
                        item.childs = [];
                    }

                    var workItem = this._queryData.filter(wi => { return wi.id == item.target.id })[0];

                    item.childs.push(workItem);
                    this.BuildTree( workItem);
                });
        }

    private ProgressMessage = function (message) {
        if (this.progressCallback != null) {
            this.progressCallback(message);
        }
    }


 }

export class BulkWIFetcher {
    private QueryData:any[] = [];
    private wiDataCount: number;
    private witClient = VSS_Service.getCollectionClient(WorkItemClient.WorkItemTrackingHttpClient); 
    
    public FetchAllWorkItems(wiIdLst, fields: any[], expand: WorkItemContracts.WorkItemExpand):IPromise<any[]> {
        var deferred = $.Deferred<any[]>();
        var reqIdLst = [];
        var size = 100;
        var asOfDate = new Date();
        
        
        console.log("No to fetch before unique : " + wiIdLst.length);
        wiIdLst = wiIdLst.unique();

        console.log("No to fetch after  unique : " + wiIdLst.length);
        this.wiDataCount = wiIdLst.length;

        this.ProgressMessage("Fetching work item data <br/>Fetched 0 out of " + this.wiDataCount);

        var bulkFetcher = this;

        if (wiIdLst.length > 0) {
            var prms: IPromise<WorkItemContracts.WorkItem[]>[] = [];
            while (wiIdLst.length > 0) {
                reqIdLst = wiIdLst.splice(0, size);
                prms.push(bulkFetcher.witClient.getWorkItems(reqIdLst, fields));
            }
            Q.all(prms).then(
                fetches => {
                    fetches.forEach(workItems => {
                        bulkFetcher.AddWorkItemData(bulkFetcher, workItems);
                    });
                    deferred.resolve(bulkFetcher.QueryData);
                },
                err => {
                    deferred.reject(err);
                });
        }
        else {
            deferred.resolve([]);  
        }

        return deferred.promise();
    }


    private AddWorkItemData(bulkFetcher: BulkWIFetcher, workItems) {

        bulkFetcher.QueryData = this.QueryData.concat(workItems);
        bulkFetcher.ProgressMessage("Fetching work item data <br/>Fetched " + this.QueryData.length + " out of " + this.wiDataCount);

        if (bulkFetcher.QueryData.length == bulkFetcher.wiDataCount) {
            return bulkFetcher.QueryData;
        }
        else {
            return null;
        }
    }

    public ProgressMessage = function (message) {
        if (this.progressCallback != null) {
            this.progressCallback(message);
        }
    }

    
}


    