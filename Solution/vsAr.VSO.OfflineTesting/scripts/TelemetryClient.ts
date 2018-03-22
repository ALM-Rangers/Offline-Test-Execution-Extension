/// <reference path="../SDK/scripts/ai.1.0.0-build00159.d.ts" />

import Context = require("VSS/Context");
import VSS_VSS = require("VSS/VSS");

export class TelemetryClient implements VSS_VSS.errorPublisher {

    private static telemetryClient: TelemetryClient;
    public static getClient(): TelemetryClient {
        if (!this.telemetryClient) {
            this.telemetryClient = new TelemetryClient();
            this.telemetryClient.Init();
        }
        return this.telemetryClient;
       
    }

    private appInsightsClient: Microsoft.ApplicationInsights.AppInsights;

    private Init() {
        var self = this;
        try {
            var snippet: any = {
                config: {
                    instrumentationKey: "__INSTRUMENTATIONKEY__"
                }
            };
            var x = VSS.getExtensionContext();
            this.appInsightsClient = null;
            //var init = new Microsoft.ApplicationInsights.Initialization(snippet);
            //this.appInsightsClient = init.loadAppInsights();

            //var webContext = VSS.getWebContext();
            //this.appInsightsClient.setAuthenticatedUserContext(
            //    webContext.user.id, webContext.collection.id);

            //window.onerror = this.appInsightsClient._onerror;
            VSS_VSS.errorHandler.attachErrorPublisher(self);

        }
        catch (e) {
            this.appInsightsClient = null;
            console.log(e);
        }
    }

    public publishError(error: TfsError): void {

        var e = new Error();
        e.name = error.name;
        e.message = error.message;
        e["stack"] = error.stack;
        if (this.appInsightsClient != null) {
            this.appInsightsClient.trackException(e)
        }
    }

    public startTrackPageView(name?: string) {
        try {
            if (this.appInsightsClient != null) {
                this.appInsightsClient.startTrackPage(name);
            }
        }
        catch (e) {
            console.log(e);
        }
    }

    public stopTrackPageView(name?: string) {
        try {
            if (this.appInsightsClient != null) {
                this.appInsightsClient.stopTrackPage(name);
            }
        }
        catch (e) {
            console.log(e);
        }
    }

    public trackPageView(name?: string, url?: string, properties?: Object, measurements?: Object, duration?: number) {
        try {
            if (this.appInsightsClient != null) {
                this.appInsightsClient.trackPageView("OfflineTesting." + name, url, properties, measurements, duration);
            }
        }
        catch (e) {
            console.log(e);
        }
    }

    public trackEvent(name: string, properties?: Object, measurements?: Object) {
        try {
            if (this.appInsightsClient != null) {
                this.appInsightsClient.trackEvent("OfflineTesting." + name, properties, measurements);
                this.appInsightsClient.flush();
            }
        }
        catch (e) {
            console.log(e);
        }
    }

    public trackException(exception: Error, handledAt?: string, properties?: Object, measurements?: Object) {
        try {
            if (this.appInsightsClient != null) {
                this.appInsightsClient.trackException(exception, handledAt, properties, measurements);
                this.appInsightsClient.flush();
            }
        }
        catch (e) {
            console.log(e);
        }
    }

    public trackMetric(name: string, average: number, sampleCount?: number, min?: number, max?: number, properties?: Object) {
        try {
            if (this.appInsightsClient != null) {
                this.appInsightsClient.trackMetric("OfflineTesting." + name, average, sampleCount, min, max, properties);
                this.appInsightsClient.flush();
            }
        }
        catch (e) {
            console.log(e);
        }
    }

}
