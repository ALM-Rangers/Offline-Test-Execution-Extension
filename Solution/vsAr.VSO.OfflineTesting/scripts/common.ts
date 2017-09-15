 
import CommonControls = require("VSS/Controls/Notifications");
import Controls = require("VSS/Controls");
 
import Treeview = require("VSS/Controls/TreeView");
import Menus = require("VSS/Controls/Menus");

import StatusIndicator = require("VSS/Controls/StatusIndicator");
import APIContracts = require("VSS/WebApi/Contracts");
import UtilsString = require("VSS/Utils/String");


export class viewBase {
    public processTemplate: string;
    public projectId: string = VSS.getWebContext().project.id;
    public tree: Treeview.TreeView;
    public nodes: Treeview.TreeNode[];
    public waitControl;

    //constructor() {
    //    var self = this;
    //    self.nodes = new Array<Treeview.TreeNode>();
    //    var home = new Treeview.TreeNode("Export Test Cases");
    //    home.link = "index.html";
    //    var sprints = new Treeview.TreeNode("Import Test Cases");
    //    sprints.link = "upload-test-suite.html";
    //    self.nodes.push(home);
    //    self.nodes.push(sprints);



    //    self.tree = Controls.create(Treeview.TreeView, $('#treeMenu'), {
    //        nodes: self.nodes

    //    });

    //   
    //    });
    //}

    public StartLoading (longRunning, message) {
        var ctrl=this;
        $("body").css("cursor", "progress");
        $("#mainDiv").css("visibility", "hidden");

        if (longRunning) {
                if (ctrl.waitControl == null) {
                    var waitControlOptions = {
                        target: $(".wait-control-target"),
                        message: message,
                        cancellable: true,
                        image:null,
                        cancelTextFormat: "{0} to cancel",
                        cancelCallback: function () {
                            console.log("cancelled");
                        }
                    };

                    VSS.require(["VSS/Controls/StatusIndicator", "VSS/Controls"], function (StatusIndicator, Controls) {
                    ctrl.waitControl = Controls.create(StatusIndicator.WaitControl, $("#previewHTML"), waitControlOptions);
                    ctrl.waitControl.startWait();
                });
            }
            else {
                ctrl.waitControl.setMessage(message);
                ctrl.waitControl.startWait();
            }
        }

    }

    public DoneLoading () {
        var ctrl=this;
        $("body").css("cursor", "default");
        $("#mainDiv")
            .css("visibility", "visible")
            .css("height","100%");

        if (ctrl.waitControl != null)
        {
            ctrl.waitControl.endWait();
        }

    }

    public ProgressUpdate (message) {
        var ctrl = this;
        if (ctrl.waitControl != null) {
            ctrl.waitControl.setMessage(message);
            return !ctrl.waitControl.isCancelled();
        }
        else{
            return true;
        }
    }
}


export interface Array<T> {
    unique(): T[];
}

Array.prototype["unique"] = function () {
    var n = {}, r = [];
    for (var i = 0; i < this.length; i++) {
        if (!n[this[i]]) {
            n[this[i]] = true;
            r.push(this[i]);
        }
    }
    return r;
}

//class UniqueArray<T> extends Array<T> {
//    public unique(): Array<T> {
//        var n = {}, r:T[] = [];
//        for (var i = 0; i < this.length; i++) {
//            if (!n[this[i]]) {
//                n[this[i]] = true;
//                r.push(this[i]);
//            }
//        }
//        return r;
//    }

//}

export interface IProgressCallback { (message: string): boolean }

export function iframeDataURITest(src): boolean {
    var support,
        iframe = document.createElement('iframe');

    iframe.style.display = 'none';
    iframe.setAttribute('src', src);

    document.body.appendChild(iframe);

    try {
        support = !!iframe.contentDocument;
    }
    catch (e) {
        support = false;
    }

    document.body.removeChild(iframe);
    return support;
}

function onCopy() {

    if (window["clipboardData"]) {
        //window["clipboardData"].setData("application/officeObj", selection.serialize());

    }
    return false;   // cancels the default copy operation

}


export function CopyContent(element) {
    var doc = document;
    var range, selection;
    element.contentEditable = true;

    if (document.body["createControlRange"]) {
        element.oncopy = onCopy;

        range = document.body["createControlRange"]();
        range.addElement(element);
        range.execCommand('Copy');

        //range.moveToElementText(element);
        // range.select();
    } else if (window.getSelection) {
        selection = window.getSelection();
        range = document.createRange();
        range.selectNodeContents(element);
        selection.removeAllRanges();
        selection.addRange(range);
        try {
            var successful = document.execCommand('copy');
            var msg = successful ? 'successful' : 'unsuccessful';
            console.log('Copying text command was ' + msg);
        } catch (err) {
            console.log('Oops, unable to copy');
        }
    }
    element.contentEditable = false;

}

export function clearSelection() {
    if (document["selection"]) {
        document["selection"].empty();
    } else if (window.getSelection) {
        window.getSelection().removeAllRanges();
    }
}


export function wrapInDiv(content) {
    return "<div>" + content + "</div>";
}


export function getHubUrl(contributionId) {
    var context = VSS.getWebContext();
    var extCont = VSS.getExtensionContext();

    return context.collection.uri + "/" + context.project.name + "/_admin/_apps/hub/" + extCont.publisherId + "." + extCont.extensionId + "." + contributionId.replace('.', '-');

}



export function TransformXml(xslText, xml) {
    var xslDocument, resultDocument, xsltProcessor, xslt, xslDoc, xslProc, transformedHtml;
    //  Diag.logVerbose("[HtmlDocumentGenerator._displayResult] Parse the xslt string");
    xslDocument = jQuery.parseXML(xslText); // Utils_Core.parseXml(xslText);

    if (window.ActiveXObject || "ActiveXObject" in window) {
        if (typeof (xml.transformNode) != "undefined") {
            transformedHtml = xml.transformNode(xslDocument);
            resultDocument = transformedHtml;
        }
        else {

            if (window.ActiveXObject || "ActiveXObject" in window) {
                xslt = new ActiveXObject("Msxml2.XSLTemplate");
                xslDoc = new ActiveXObject("Msxml2.FreeThreadedDOMDocument");
                xslDoc.loadXML(xslText);
                xslt.stylesheet = xslDoc;
                xslProc = xslt.createProcessor();
                xslProc.input = xml;
                xslProc.transform();
                transformedHtml = xslProc.output;
                resultDocument = transformedHtml;
            }

        }
    }
    else if (document.implementation && document.implementation.createDocument) {
        xsltProcessor = new window.XSLTProcessor();
        xsltProcessor.importStylesheet(xslDocument);
        transformedHtml = xsltProcessor.transformToDocument(xml);
        resultDocument = transformedHtml.documentElement.outerHTML;
    }
    return resultDocument;
};

export function xmlEscape(s) {
    if ($.type(s) === "string") {
        return (s
            .replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/'/g, '&apos;')
            .replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\t/g, '&#x9;').replace(/\n/g, '&#xA;').replace(/\r/g, '&#xD;')
        );
    }
    else {
        return s;
    }
}



/**
 * XML2jsobj v1.0
 * Converts XML to a JavaScript object
 * so it can be handled like a JSON message
 *
 * By Craig Buckler, @craigbuckler, http://optimalworks.net
 *
 * As featured on SitePoint.com:
 * http://www.sitepoint.com/xml-to-javascript-object/
 *
 * Please use as you wish at your own risk.
 */

export function XML2jsobj(node) {

    var data = {};

    // append a value
    function Add(name, value) {
        if (data[name]) {
            if (data[name].constructor != Array) {
                data[name] = [data[name]];
            }
            data[name][data[name].length] = value;
        }
        else {
            data[name] = value;
        }
    };

    // element attributes
    var c, cn;
    for (c = 0; cn = node.attributes[c]; c++) {
        Add(cn.name, cn.value);
    }

    // child elements
    for (c = 0; cn = node.childNodes[c]; c++) {
        if (cn.nodeType == 1) {
            if (cn.childNodes.length == 1 && cn.firstChild.nodeType == 3) {
                // text value
                Add(cn.nodeName, cn.firstChild.nodeValue);
            }
            else {
                // sub-object
                Add(cn.nodeName, XML2jsobj(cn));
            }
        }
    }

    return data;

}




export interface HTMLFileElement extends HTMLElement {
    files: FileList;
}

declare var FileReader: {
    new ();
    readAsBinaryString(f);
}


export function SaveFile(data, fileName, type) {
    if (window.navigator.msSaveOrOpenBlob != null) {
        SaveFileMsBlob(str2bytes(data), fileName, type);
    }
    else {
        SaveFileDataUri(
            btoa(data),
            fileName,
            type);
    }
}


export function SaveFileDataUri(data, fileName, type) {
    var a: HTMLAnchorElement = <HTMLAnchorElement>document.body.appendChild(
        document.createElement("a")
    );

    a["download"] = fileName;
    a.href = "data:text/" + type + ";base64," + data;
    a.innerHTML = "download";
    a.click();
    document.body.removeChild(a);
    //  delete a;
}

export function SaveFileMsBlob(data, fileName, mtype) {
    var blob = new Blob([data], { type: mtype });  //"data:application/zip;base64"});

    window.navigator.msSaveOrOpenBlob(blob, fileName);
}

export function s2ab(s) {
    var buf = new ArrayBuffer(s.length);
    var view = new Uint8Array(buf);
    for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
    return buf;
}

function str2bytes(str) {
    var bytes = new Uint8Array(str.length);
    for (var i = 0; i < str.length; i++) {
        bytes[i] = str.charCodeAt(i);
    }
    return bytes;
}

export function secureForFileName(s:string) {
    s = s.replace(/[^[:alnum:]] /gi, "");
    console.log(s);
    return s;
}

export function getIconFromTestOutcome(outcome): string {
    var icon: string = "";
    switch (outcome) {
        case "NotApplicable":
            icon = "icon-tfs-tcm-not-applicable";
            break;
        case "Blocked":
            icon = "icon-tfs-tcm-block-test";
            break;
        case "Passed":
            icon = "icon-tfs-build-status-succeeded";
            break;
        case "Failed":
            icon = "icon-tfs-build-status-failed";
            break;
        case "None":
            icon = "icon-tfs-tcm-block-test";
            break;
        case "DynamicTestSuite":
            icon = "icon-tfs-build-status-succeeded";
            break
    }
    return icon;
}


export function getDisplayName(i: APIContracts.IdentityRef): string {
    return i.displayName != null ? i.displayName.split("<")[0] : "";
}