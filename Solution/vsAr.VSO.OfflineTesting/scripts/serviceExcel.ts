/// <reference types="vss-web-extension-sdk" />

var datenum;
import Controls = require("VSS/Controls");
import FileInput = require("VSS/Controls/FileInput");
import StatusIndicator = require("VSS/Controls/StatusIndicator");

import Common = require("scripts/Common");
import TestHelper = require("scripts/TestHelper");

var strComment = "Comment";

export interface IImportData {
    testCaseId: number,
    title: string,
    testPointId: number,
    iteration?: number,
    config: string,
    tester: string,
    outcome?: string
    comment?: string,
    execDate?: Date,
    steps?: TestHelper.ITestStep[],
    runBy?: any,
    validErrMsg?: string,
    validWarnMsg?: string,
    parameters?: { [key: string]: string },
}



export function importFromExcel(input): { plan: string, data: IImportData[], minDate: Date, maxDate: Date} {
    var self = this;
    var testResults: IImportData[] = [];
    var el = <FileInput.FileInputControl>input;
    var files = el.getFiles();
    var file = files[0];

    
    var workbook = XLSX.read(file.content, { type: 'base64' });
    /* We are only concerned with the first worksheet for now */

    var worksheet = workbook.Sheets[workbook.SheetNames[0]];
    var src = XLSX.utils.sheet_to_json(worksheet);
    var minDt: Date = null;
    var maxDt: Date = null;
    src.filter((x:any)=> { return x.TestPointId != '' }).forEach((itm: any) => {
        //"TestCaseId", "Title", "TestPointId", "Configuration", "Tester", "TestStep", "StepAction", "StepExpected",  "Outcome", "Comment"]
        if (itm.TestPointId != null) {
            var testpointParts = itm.TestPointId.split(":");
            var iteration = testpointParts.length > 1 ? testpointParts[1] : null;
            console.log(testpointParts);

            var r: IImportData = {
                testCaseId: Number(itm.TestCaseId),
                title: itm.Title,
                testPointId: testpointParts[0],
                iteration: iteration,
                config: itm.Configuration,
                tester: itm.Tester,
                outcome: itm.Outcome,
                //execDate: (itm.Date !="" && itm.Date!= null)?new Date(itm.Date):null,                    
                comment: itm[strComment],
                parameters: iteration != null ? {} : null
            }

            if (r.execDate != null) {
                if (r.execDate < minDt || minDt == null) {
                    minDt = r.execDate;
                }
                if (r.execDate > maxDt || maxDt == null) {
                    maxDt = r.execDate;
                }
            }

            var rowNum = itm.__rowNum__;
            while (rowNum < src.length && rowNum != -1) {
                var stepRow: any = src[rowNum];
                if (stepRow.TestPointId == '') {
                    if (r.steps == null) {
                        r.steps = [];
                    }
                    if (stepRow.TestStepId != null) {
                        //Old export contained TestStepId
                        r.steps.push({ id: stepRow.TestStepId, index: stepRow.TestStepId, action: stepRow.StepAction, expectedResult: stepRow.StepExpected, outcome: stepRow.Outcome, comment: stepRow[strComment], isFormatted: false });
                    }
                    else {
                        r.steps.push({ index: stepRow.TestStep, action: stepRow.StepAction, expectedResult: stepRow.StepExpected, outcome: stepRow.Outcome, comment: stepRow[strComment], isFormatted: false });
                    }

                    if (iteration != null) {
                        ScanStringForParameters(stepRow.StepAction, r.parameters);
                        ScanStringForParameters(stepRow.StepExpected, r.parameters);
                    }

                    rowNum++;
                }
                else {
                    rowNum = -1;
                }
            }
            var hasStepsOutcome = (r.steps != null && r.steps.filter(s => { return s.outcome != ""; }).length > 0);
            if ((itm.Outcome != null && itm.Outcome != "") || hasStepsOutcome) {
                testResults.push(r);
            }
        }
        else {
            console.log("No testpoint returned ");
            console.log(itm);
        } 


    });
    return { plan: workbook.SheetNames[0],data: testResults, minDate:minDt, maxDate:maxDt}

}

export function exportToExcel(sheetName, tpList: IImportData[], opts):any {
    var xlWB = createExcel(sheetName, tpList, opts);
    return xlWB;
}

function createExcel(sheetName, tpList:IImportData[], opts):any {
    var ws = {};
    // Convert objects to two dimensional array
    var data: Array<Array<any>> = new Array<Array<any>>();

    data[0] = ["TestCaseId", "Title", "TestStep", "StepAction", "StepExpected", "TestPointId", "Configuration", "Tester", "Outcome", strComment];
    var rowCounter = 1;
    function getRowData(test: IImportData) {
        var testPointStr = $.isNumeric(test.iteration) ? test.testPointId + ":" + test.iteration: test.testPointId ;
        data[rowCounter] = [test.testCaseId, test.title, "", "", "", testPointStr, test.config, test.tester, "",  ""];
        rowCounter++;
           
        if (test.steps != undefined) {
            test.steps.forEach((step ,ix)=> {
                var action = step.isFormatted ? unFormat(step.action) : step.action;
                var result = step.isFormatted ? unFormat(step.expectedResult) : step.expectedResult;
                console.log("   extracting step " + step.id + " (ix:" + step.index + ")" + action + ":" + result);
                
                data[rowCounter] = ["", "", step.index, action, result, "", "", "", "",  ""];
                rowCounter++;
            });
        }
    }
    tpList.forEach(i => {
          
            getRowData(i);
    });

    // format the 2D array as worksheet cells
    var range = { s: { c: 10000000, r: 10000000 }, e: { c: 0, r: 0 } };
    for (var R = 0; R != data.length; ++R) {
        for (var C = 0; C != data[R].length; ++C) {
            if (range.s.r> R) range.s.r = R;
            if (range.s.c > C) range.s.c = C;
            if (range.e.r < R) range.e.r = R;
            if (range.e.c < C) range.e.c = C;
            var cell: XLSX.IWorkSheetCell = { v: data[R][C], t: 's' };

            if (cell.v == null) continue;
            var cell_ref = XLSX.utils.encode_cell({ c: C, r: R });

            if (typeof cell.v === 'number') cell.t = 'n';
            else if (typeof cell.v === 'boolean') cell.t = 'b';
            else if (<any>cell.v instanceof Date) {
                cell.t = 'n';
                cell.z = XLSX.SSF._table[14];
                cell.v = datenum(cell.v);
            }
            else {
                cell.t = 's';
                cell.s = '1';
            }
            
            ws[cell_ref] = cell;
        }
    }

    if (range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
    /* bookType can be 'xlsx' or 'xlsm' or 'xlsb' */
    var workbook: any = {
        SheetNames: [],
        Sheets: {}
    }

    workbook.SheetNames.push(Common.xmlEscape(sheetName));
    ws['!dataValidations'] = [{ type: "list", allowBlank: "1", showInputMessage: "1", showErrorMessage: "1", range: "I1:I1048576", formula: "Passed,Failed,Blocked,Paused" }];
    ws['!cols'] = [
        { wch: 10 },
        { wch: 10 },
        { wch: 10 },
        { wch: 30 },
        { wch: 30 },
        { wch: 10 },
        { wch: 12 },
        { wch: 15 },
        { wch: 10 },
        { wch: 10 }

    ];

    
    workbook.Sheets[sheetName] = ws;
    
    var wbout = XLSX.write(workbook, opts);
    return wbout
}

function ScanStringForParameters(txt : string, params: {[key: string]: string }):void{
    var regex = /(\[@)([^=]+)( = )([^=]+)(\])/g

    var matches = [];
    var match = regex.exec(txt);
    while (match != null) {
        params[match[2]]=match[4]
        match = regex.exec(txt);
    }
}




function unwind(elem) {

    if (elem.P !== undefined) {
        return unwind(elem.P);

    }
    else if (elem.div !== undefined) {
        return unwind(elem.div);

    }
    else if (elem.DIV !== undefined) {
        return unwind(elem.DIV);
    }
    else {

        if (elem.toString() != "[object Object]") {
            return elem.replace("\n", "");
        }
        else {
            return "";
        }

    }
}


function unFormat(html) {
    var htElem = $.parseHTML(html);
    var s = "";//<r><t>";
    if (html.length > 0) {
        htElem.forEach(function (e) {

            //if (e.innerHTML != null) {
            //    var frmt = e.innerHTML;
            //    frmt = frmt.replace(new RegExp("<br>", "gi"), "\n");
            //    frmt = frmt.replace(new RegExp("&nbsp;", "gi"), " ");
            //    frmt = frmt.replace(new RegExp("<b>", "gi"), "<r><rPr><b/></rPr><t>");
            //    frmt = frmt.replace(new RegExp("</b>", "gi"), "</t></r>");
            //    frmt = frmt.replace(new RegExp("<i>", "gi"), "<r><rPr><i/></rPr><t>");
            //    frmt = frmt.replace(new RegExp("</i>", "gi"), "</t></r>");
            //    s = s + frmt + "\n";

            //} else

            if (e.innerText != null) {
                s = s + e.innerText + "\n";
            } else if (e.nodeValue != null) {
                s = s + e.nodeValue + "\n";
            }
        });
    }
    return s;
    //   return "<si><r><t xml:space = 'preserve' > Step 1 _x000D__x000D_Multi line with </t></r><r> <rPr><b/><sz val='12'/> <color theme='1' /><rFont val='Calibri' /><family val='2' /><scheme val='minor' /></rPr><t>BOLD</t> </r><r><rPr><sz val='12'/> <color theme='1' /><rFont val='Calibri' /><family val='2' /><scheme val='minor' /></rPr><t xml:space='preserve'> , </t> </r><r><rPr><i/> <sz val='12' /><color theme='1' /><rFont val='Calibri' /><family val='2' /><scheme val='minor' /></rPr><t>italics</t> </r><r><rPr><sz val='12'/> <color theme='1' /><rFont val='Calibri' /><family val='2' /><scheme val='minor' /></rPr><t xml:space='preserve'> and a_x000D_</t></r> <r><rPr><sz val='10' /><color theme='1' /><rFont val='Calibri' /><family val='2' /><scheme val='minor' /></rPr><t>list _x000D_ of_x000D_points </t></r> <r><rPr><sz val='12' /><color theme='1' /><rFont val='Calibri' /><family val='2' /><scheme val='minor' /></rPr><t xml:space='preserve'></t></r> </si>";
}
