/**
 * Prompt an alert message
 * 
 * @param {string} message Alert message text
 */
export function alertMessage(message: string): void{
    (Xrm.Utility !== undefined && Xrm.Utility.alertDialog !== undefined) ? (<any>Xrm).Utility.alertDialog(message) : alert(message);
}

/**
 * Check if two guids are equal
 * 
 * @export
 * @param {string} guid1 A string represents a guid
 * @param {string} guid2 A string represents a guid
 * @returns {boolean}
 */
export function guidsAreEqual(guid1: string, guid2: string): boolean{
    let isEqual: boolean;
    if (guid1 === null || guid2 === null || guid1 === undefined || guid2 === undefined) {
        isEqual = false;
    } else {
        isEqual = guid1.replace(/[{}]/g, "").toLowerCase() === guid2.replace(/[{}]/g, "").toLowerCase();
    }

    return isEqual;
}

/**
 * Private function to the context object.
 * 
 * @export
 * @returns {Xrm.Context}
 */
export function context(): Xrm.Context {
    let oContext: Xrm.Context;
    if (typeof window.GetGlobalContext !== "undefined") {
        oContext = window.GetGlobalContext();
    } else if (typeof GetGlobalContext !== "undefined") {
        oContext = GetGlobalContext();
    } else {
        if (typeof Xrm !== "undefined") {
            oContext = Xrm.Page.context;
        } else if (typeof window.parent.Xrm !== "undefined") {
            oContext = window.parent.Xrm.Page.context;
        } else {
            throw new Error("Context is not available.");
        }
    }
    return oContext;
}

 /**
  * Private function to return the server URL from the context
  * 
  * @export
  * @returns {string} Url of the organization
  */
 export function getClientUrl() {
    let serverUrl = typeof context().getClientUrl !== "undefined" ? context().getClientUrl() : (<any>context()).getServerUrl();
    if (serverUrl.match(/\/$/)) {
        serverUrl = serverUrl.substring(0, serverUrl.length - 1);
    }
    return serverUrl;
}