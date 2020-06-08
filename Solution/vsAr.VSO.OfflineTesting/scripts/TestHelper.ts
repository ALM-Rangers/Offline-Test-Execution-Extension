

import WorkItemContracts = require("TFS/WorkItemTracking/Contracts");
import UtilsString = require("VSS/Utils/String");

import Telemetry = require("scripts/TelemetryClient");

export interface IParam { type: string, sharedParamId: number, title: string, params: any[] };

export class TestOucomeAggregator {
    protected prio: { [keyof: string]: number };
    protected totOutComePrio: number;
    protected outcomeCnt: number;
    constructor() {
        this.prio = {};
        
        this.prio["Passed"] = 1;
        this.prio["Active"] = 2;
        this.prio["Paused "] = 3;
        this.prio["Failed"] = 4;
        this.prio["Blocked"] = 5;
        this.totOutComePrio = 0;
        this.outcomeCnt = 0;
    }

    public addOutcome(outcome: string) {
        this.totOutComePrio = Math.max(this.totOutComePrio, this.prio[outcome]);
        this.outcomeCnt++;
    }

    public getOutcome() {
        var outcome = "Passed";

        for (var s in this.prio) {
            if (this.prio.hasOwnProperty(s)) {
                if (this.prio[s] === this.totOutComePrio) {
                    outcome = s;
                    break;
                }
            }
        }
        return outcome;
    }
    public getOutcomeCount() {
        return this.outcomeCnt;
    }
}



export interface ITestStepAttachments {
    testStepId?: number,
    id?: number,
    size: number;
    name: string,
    url: string;
}

export interface ITestStep {
    id?: number,
    index: string,
    stepType?: string,
    action: string,
    expectedResult: string,
    isFormatted: boolean,
    outcome?: string,
    comment?: string
    attachments?: ITestStepAttachments[],
    sharedStepWorkItemId?: number,
    sharedStepWorkItemRevision?: number
    parentStepId?: number;
    params?: { [key: string]: string };
};


export function scanSharedStepsAndParameters(listTc: WorkItemContracts.WorkItem[], scanTestSteps: boolean, scanParameters: boolean): Number[] {
    {
        var sharedIds: number[] = [];

        var self = this;
        listTc.forEach(tc => {
            if (scanTestSteps) {
                if (tc.fields["Microsoft.VSTS.TCM.Steps"] != null) {

                    var $xml = $.parseXML(tc.fields["Microsoft.VSTS.TCM.Steps"]);

                    var lst = $xml.querySelectorAll("compref");
                    if (lst != null) {
                        for (var i = 0; i < lst.length; i++) {
                            sharedIds.push(Number(lst[i].getAttribute("ref")));
                        }
                    }
                }
            }
            if (scanParameters) {
                if (tc.fields["Microsoft.VSTS.TCM.LocalDataSource"] != null) {
                    var json = tc.fields["Microsoft.VSTS.TCM.LocalDataSource"];
                    if (json.charAt(0) == "{") {
                        var params = JSON.parse(json);
                        var sharedDataSetIds = params.sharedParameterDataSetIds;
                        sharedDataSetIds.forEach(id => {
                            sharedIds.push(id);
                        });
                    }
                }

            }
        });
        return sharedIds;
    }
}

export class TestHelper {
    protected preFetchedSharedItems: WorkItemContracts.WorkItem[];

    public setSharedWIData(data: WorkItemContracts.WorkItem[]) {
        this.preFetchedSharedItems = data;
    }

    public parseTestCaseSteps(tcWi: WorkItemContracts.WorkItem, params?: any): ITestStep[] {
        var stepsXml = tcWi.fields['Microsoft.VSTS.TCM.Steps'];
        var attatchments = this.scanTestStepAttachments(tcWi);

        var steps: ITestStep[] = [];
        try {
            if (stepsXml != null) {
                console.log(stepsXml);
                var xml = $.parseXML(stepsXml);
                var stepListRoot: Node = xml.childNodes[0];
                if (stepListRoot.nodeName != "steps") {
                    stepListRoot = xml.getElementsByTagName("steps")[0];
                    Telemetry.TelemetryClient.getClient().trackEvent("Unexpected xml structure", { xml: stepsXml });
                }
                steps = this.parseTestSteps(stepListRoot.childNodes, attatchments, "", 0);

                if (params != null) {
                    steps.forEach(s => {
                        for (var p in params) {
                            if (params.hasOwnProperty(p)) {
                                var regexp = new RegExp("@" + p, 'gi')
                                if (regexp.test(s.action) || regexp.test(s.expectedResult)) {
                                    if (s.params == null) {
                                        s.params = {};
                                    }
                                    console.log("Setting " + p + "=" + params[p]);
                                    s.params[p] = params[p];
                                    s.parentStepId
                                    s.action = s.action.replace(regexp, "[@" + p + " = " + params[p] + "]");
                                    s.expectedResult = s.expectedResult.replace(regexp, "[@" + p + " = " + params[p] + "]");

                                    console.log(s.action + " =====> " + s.expectedResult);
                                }
                            }
                        }
                    });
                }
            }
        } catch (ex) {
            var prop = {};
            stepsXml.split("<step").forEach((s, ix) => {
                prop["Step_" + ix] = s;
            });

            Telemetry.TelemetryClient.getClient().trackException(ex, "", prop);
            return [];
        }
        return steps;
    }

    protected scanTestStepAttachments(tcWI: WorkItemContracts.WorkItem): ITestStepAttachments[] {
        try {
            var regexp = /(\[TestStep=)(\d+)(\]:)/i;
            if (tcWI && tcWI.relations) {
                var attatchments = tcWI.relations.filter(i => {
                    return i.rel == "AttachedFile"
                        && i.attributes["comment"] != null
                        && i.attributes["comment"].indexOf("[TestStep=") >= 0
                });
                return attatchments.map(r => {
                    var testStepId = r.attributes["comment"].match(regexp);

                    return {
                        testStepId: testStepId.lenght = 4 ? testStepId[2] : -1,
                        id: r.attributes["id"],
                        size: r.attributes["resourceSize"],
                        name: r.attributes["name"],
                        url: r.url
                    }
                });
            }
            else {
                return [];
            }
        }
        catch (ex) {
            Telemetry.TelemetryClient.getClient().trackException(ex, "", { relations: tcWI.relations });
            return [];
        }
    }

    protected parseTestSteps($stepList: NodeList, attachments: ITestStepAttachments[], index: string, startIx: number): ITestStep[] {

        var self = this;
        var steps: ITestStep[] = [];
        var stepIx = 0;
        for (var i = 0; i < $stepList.length; i++) {
            var step = $stepList[i];
            console.log(step);
            switch (step.nodeName) {
                case "step":
                    stepIx++;
                    var stp = readStep(step, index + (stepIx + startIx));
                    var lst = attachments.filter(i => { return i.testStepId == stp.id });
                    if (lst.length > 0) {
                        stp.attachments = lst;
                        console.log("***** " + stp.attachments.length);
                        console.log(stp.attachments);
                    }

                    steps.push(stp);
                    break;
                case "compref":
                    stepIx++;
                    steps = steps.concat(self.parseSharedSteps(step, attachments, index, stepIx, startIx));
                    break;
                case "#text":
                    var ex = new Error();
                    ex.message = "Unexpected #text step data at parseTestSteps";
                    ex.name = "Unexpected# text step ";
                    var str = step.textContent
                    Telemetry.TelemetryClient.getClient().trackException(ex, "", { nodeName: step.nodeName, outerXml: str, value: step.nodeValue });

                    break;
                default:
                    var ex = new Error();
                    ex.message = "Unexpected step data at parseTestSteps";
                    ex.name = "Unexpected step ";
                    var str = step.textContent;
                    Telemetry.TelemetryClient.getClient().trackException(ex, "", { nodeName: step.nodeName, outerXml: str, value: step["innerHTML"] });
                    break;
            }
        }
        console.log(steps);
        return steps;
    }

    protected parseSharedSteps(step: Node, attachments: ITestStepAttachments[], index: string, i: number, startIx: number): ITestStep[] {
        var self = this;
        var sharedStepWorkItemId = Number(step.attributes.getNamedItem("ref").value);
        var parentId = parseInt($(step).attr("id"), 10);

        var steps: ITestStep[] = [];
        if (self.preFetchedSharedItems) {
            var sharedStepWI = self.preFetchedSharedItems.filter(i => { return i.id == sharedStepWorkItemId; });

            if (sharedStepWI.length > 0) {
                var sharedStepAttatchments = this.scanTestStepAttachments(sharedStepWI[0]);
                var stepDataXml = sharedStepWI[0].fields['Microsoft.VSTS.TCM.Steps'];

                try {
                    var sharedIx = index + (i + startIx) + "."
                    steps.push({
                        id: parentId,
                        index: index + (i + startIx),
                        action: sharedStepWI[0].fields['System.Title'],
                        expectedResult: "",
                        isFormatted: false,
                        sharedStepWorkItemId: sharedStepWorkItemId,
                        sharedStepWorkItemRevision: sharedStepWI[0].rev
                    });

                    var xml = $.parseXML(stepDataXml);
                    var stepListRoot: Node = xml.childNodes[0];
                    if (stepListRoot.nodeName != "steps") {
                        stepListRoot = xml.getElementsByTagName("steps")[0];
                        Telemetry.TelemetryClient.getClient().trackEvent("Unexpected xml structure", { xml: stepDataXml });
                    }

                    var sharedSteps = self.parseTestSteps(stepListRoot.childNodes, sharedStepAttatchments, sharedIx, 0);
                    steps = steps.concat(sharedSteps.map(o => {
                        o.parentStepId = parentId;
                        //o.sharedStepWorkItemId = sharedStepWorkItemId;
                        //o.sharedStepWorkItemRevision = sharedStepData[0].rev;
                        return o;
                    }));

                    if (step.childNodes != null) {
                        steps = steps.concat(self.parseTestSteps(step.childNodes, attachments, index, i + startIx));
                    }
                }
                catch (ex) {
                    var prop = {};
                    if (stepDataXml != null) {
                        stepDataXml.split("<step").forEach((s, ix) => {
                            prop["Step_" + ix] = s;
                        });
                    }
                    else {
                        prop["stepData"] = stepDataXml;
                    }
                    console.log(prop);

                    Telemetry.TelemetryClient.getClient().trackException(ex, "", prop);
                    return [];
                }
            }
        }
        else {
            steps.push({
                id: sharedStepWorkItemId,
                index: index + (i + startIx),
                action: "ERROR MISSING SHARED STEP WorkItemId=" + sharedStepWorkItemId,
                expectedResult: "",
                isFormatted: false
            });

        }

        return steps;
    }

    public parseTestParams(paramData: any): IParam {
        var self = this;
        var returData: IParam = {
            type: "",
            sharedParamId: -1,
            title: "",
            params: null
        };
        var iterations = [];
        // var self = this;
        var params = parseJSON(paramData);

        if (params != null) {
            returData.type = "SHARED";
            try {
                var sharedDataSetIds = params.sharedParameterDataSetIds;
                returData.sharedParamId = sharedDataSetIds[0];

                var sharedParamData = self.preFetchedSharedItems.filter(i => { return i.id == returData.sharedParamId; });
                if (sharedParamData.length > 0) {
                    console.log("matched prefetched")
                    var shrdParamData = sharedParamData[0].fields['Microsoft.VSTS.TCM.Parameters'];
                    returData.title = sharedParamData[0].fields['System.Title'];

                    returData.params = parseSharedParamXml(shrdParamData)
                }
            } catch (ex) {
                console.log(paramData);
                Telemetry.TelemetryClient.getClient().trackException(ex, "parseTestParams", { data: paramData });
            }

        }
        else {
            returData.type = "LOCAL";
            returData.params = parseParamXml(paramData);
        }

        return returData;
    }

}


function parseJSON(data: any): any {
    var r: any = null;
    try {
        var o = JSON.parse(data);
        if (o && typeof o === "object") {
            r = o;
        }

    }
    catch (ex) {

    }
    return r
}





function parseParamXml(paramData): any[] {

    var $xmlDom = $.parseXML(paramData);
    var iterations = [];
    $($xmlDom).find("Table1").each(function (i, row) {
        if (row.childNodes != null) {
            var it = {};
            for (var i = 0; i < row.childNodes.length; i++) {
                var param = row.childNodes[i];
                it[param.nodeName] = param.textContent;
            }
            iterations.push(it);
        }
    });

    return iterations;
}

function parseSharedParamXml(paramData): any[] {

    var $xmlDom = $.parseXML(paramData);
    var iterations = [];
    $($xmlDom).find("dataRow ").each(function (i, row) {
        if (row.childNodes != null) {
            var it = {};
            for (var i = 0; i < row.childNodes.length; i++) {
                var param = row.childNodes[i];
                if (param.attributes != null) {
                    it[param.attributes.getNamedItem("key").value] = param.attributes.getNamedItem("value").value;
                }
            }
            iterations.push(it);
        }
    });

    return iterations;
}

function readStep(step, index: string): ITestStep {
    var id = $(step).attr("id");
    var stepType = $(step).attr("type");
    var action: string, expectedResult: string;
    var count = 0, isFormatted = false, isActionFormatted = false, isExpectedResultFormatted = false;

    $(step).children("parameterizedString").each(function () {
        if (count === 0) {
            action = readParameterizedString(this);
            isActionFormatted = ($(this).attr("isformatted") === "true");
        }
        else {
            expectedResult = readParameterizedString(this);
            isExpectedResultFormatted = ($(this).attr("isformatted") === "true");
        }
        count++;
    });
    if (isActionFormatted && isExpectedResultFormatted) {
        isFormatted = true;
    }
    if (!action && !expectedResult) {
        $(step).children().each(function () {
            switch (this.nodeName.toLowerCase()) {
                case "action":
                    action = $(this).text();
                    break;
                case "expected":
                    expectedResult = $(this).text();
                    break;
            }
        });
    }

    var testStep: ITestStep = {
        id: parseInt(id, 10),
        index: index,
        stepType: stepType,
        action: action,
        expectedResult: expectedResult,
        isFormatted: isFormatted
    };

    return testStep;
}

function readParameterizedString(parameterizedString) {
    var stringParts = [];
    if ($(parameterizedString).children().length > 0) {
        $(parameterizedString).children().each(function () {
            switch (this.nodeName.toLowerCase()) {
                case "parameter":
                    stringParts.push(UtilsString.format("@{0}", $(this).text()));
                    break;
                case "outputparameter":
                    stringParts.push(UtilsString.format("@?{0}", $(this).text()));
                    break;
                case "text":
                    stringParts.push($(this).text());
                    break;
            }
        });
    }
    else {
        stringParts.push($(parameterizedString).text());
    }
    return stringParts.join("");
}
