import {getClientUrl, alertMessage} from "./Helper";

export function htmlEncode(s: string): string {
    let buffer: string = "";
    let hEncode: string = "";
    if (s === null || s === "" || s === undefined) return s;
    for (let count = 0, cnt = 0, slength = s.length; cnt < slength; cnt++) {
        let c = s.charCodeAt(cnt);
        if (c > 96 && c < 123 || c > 64 && c < 91 || c === 32 || c > 47 && c < 58 || c === 46 || c === 44 || c === 45 || c === 95){
            buffer += String.fromCharCode(c);
        } else {
            buffer += "&#" + c + ";";
        }
        if (++count === 500) {
            hEncode += buffer; buffer = ""; count = 0;
        }
    }
    if (buffer.length) hEncode += buffer;
    return hEncode;
}

export function innerSurrogateAmpersandWorkaround(s: string): string {
    let buffer: string = "";
    let c0: number;
    let cnt: number = 0;
    let slength: number = s.length;
    for ( ; cnt < slength; cnt++) {
        c0 = s.charCodeAt(cnt);
        if (c0 >= 55296 && c0 <= 57343) {
            if (cnt + 1 < s.length) {
                let c1 = s.charCodeAt(cnt + 1);
                if (c1 >= 56320 && c1 <= 57343) {
                    buffer += "CRMEntityReferenceOpen" + ((c0 - 55296) * 1024 + (c1 & 1023) + 65536).toString(16) + "CRMEntityReferenceClose"; cnt++;
                } else {
                    buffer += String.fromCharCode(c0);
                }
            } else {
                buffer += String.fromCharCode(c0);
            }
        } else {
            buffer += String.fromCharCode(c0);
        }
    }
    s = buffer;
    buffer = "";
    for (cnt = 0, slength = s.length; cnt < slength; cnt++) {
        c0 = s.charCodeAt(cnt);
        if (c0 >= 55296 && c0 <= 57343){
            buffer += String.fromCharCode(65533);
        } else {
            buffer += String.fromCharCode(c0);
        }
    }
    s = buffer;
    s = htmlEncode(s);
    s = s.replace(/CRMEntityReferenceOpen/g, "&#x");
    s = s.replace(/CRMEntityReferenceClose/g, ";");
    return s;
}

export function crmXmlEncode(s: string): string {
    if ("undefined" === typeof s || "unknown" === typeof s || null === s){
         return s;
    } else if (typeof s !== "string") {
         s = s.toString();
        }
    return innerSurrogateAmpersandWorkaround(s);
}

export function crmXmlDecode(s: string): string {
    if (typeof s !== "string") {
        s = s.toString();
    }
    return s;
}

/**
 * Private function to return the path to the REST endpoint.
 * 
 * @export
 * @returns String of the OrganizationData Service
 */
export function oDataPath() {
    return getClientUrl() + "/XRMServices/2011/OrganizationData.svc/";
}

/**
 * Private function return an Error object to the errorCallback
 * 
 * @export
 * @param {XMLHttpRequest} req The XMLHttpRequest response that returned an error.
 */
export function errorHandler(req: XMLHttpRequest) {
    throw new Error("Error : " +
    req.status + ": " +
    req.statusText + ": " +
    JSON.parse(req.responseText).error.message.value);
}

/**
 * Private function to convert matching string values to Date objects.
 * 
 * @export
 * @param {string} key The key used to identify the object property
 * @param {string} value The string value representing a date
 * @returns {(string | Date)}
 */
export function dateReviver(key: string, value: string): string | Date {
    let a: RegExpExecArray;
    if (typeof value === "string") {
        a = /Date\(([-+]?\d+)\)/.exec(value);
        if (a) {
            return new Date(parseInt(value.replace("/Date(", "").replace(")/", ""), 10));
        }
    }
    return value;
}

/**
 * Private function used to check whether required parameters are null or undefined
 * 
 * @export
 * @param parameter The parameter to check
 * @param message The error message text to include when the error is thrown
 */
export function parameterCheck(parameter: any, message: string): void | Error {
    if ((typeof parameter === "undefined") || parameter === null) {
        throw new Error(message);
    }
}

/**
 * Private function used to check whether required parameters are null or undefined
 * 
 * @export
 * @param {*} parameter The string parameter to check
 * @param {string} message The error message text to include when the error is thrown
 */
export function stringParameterCheck(parameter: any, message: string): void | Error {
    if (typeof parameter !== "string") {
        throw new Error(message);
    }
}

/**
 * Private function used to check whether required callback parameters are functions
 * 
 * @export
 * @param {*} callbackParameter The callback parameter to check
 * @param {string} message The error message text to include when the error is thrown.
 */
export function callbackParameterCheck(callbackParameter: any, message: string): void | Error {
    if (typeof callbackParameter !== "function") {
        throw new Error(message);
    }
}

/**
 * Private function used to check whether required parameters are null or undefined
 * 
 * @export
 * @param {*} parameter The boolean parameter to check
 * @param {string} message The error message text to include when the error is thrown
 */
export function booleanParameterCheck(parameter: any, message: string): void | Error {
    if (typeof parameter !== "boolean") {
        throw new Error(message);
    }
}

/**
 * Get an instance of XMLHttpRequest for all browsers
 * 
 * @export
 * @returns (description)
 */
export function getXhr(): XMLHttpRequest | ActiveXObject {
    if (XMLHttpRequest) {
        // Chrome, Firefox, IE7+, Opera, Safari
        // ReSharper disable InconsistentNaming
        return new XMLHttpRequest();
        // ReSharper     restore InconsistentNaming
    }
    // IE6
    try {
        // The latest stable version. It has the best security, performance,
        // reliability, and W3C conformance. Ships with Vista, and available
        // with other OS's via downloads and updates.
        return new ActiveXObject("MSXML2.XMLHTTP.6.0");
    } catch (e) {
        try {
            // The fallback.
            return new ActiveXObject("MSXML2.XMLHTTP.3.0");
        } catch (e) {
            alertMessage("This browser is not AJAX enabled.");
            return null;
        }
    }
}

export function performRequest(settings) {
    parameterCheck(settings, "The value passed to the performRequest function settings parameter is null or undefined.");
    var request = getXhr();
    request.open(settings.type, settings.url, settings.async);
    request.setRequestHeader("Accept", "application/json");
    if (settings.action != null) {
        request.setRequestHeader("X-HTTP-Method", settings.action);
    }
    request.setRequestHeader("Content-Type", "application/json; charset=utf-8");

    if (settings.async) {
        request.onreadystatechange = function () {
            if (request.readyState === 4 /*Complete*/) {
                // Status 201 is for create, status 204/1223 for link and delete.
                // There appears to be an issue where IE maps the 204 status to 1223
                // when no content is returned.
                if (request.status === 204 || request.status === 1223 || request.status === 201) {
                    settings.success(request);
                }
                else {
                    // Failure
                    if (settings.error) {
                        settings.error(errorHandler(request));
                    }
                    else {
                        errorHandler(request);
                    }
                }
            }
        };

        if (typeof settings.data === "undefined") {
            request.send();
        }
        else {
            request.send(settings.data);
        }
    } else {
        if (typeof settings.data === "undefined") {
            request.send();
        }
        else {
            request.send(settings.data);
        }
        if (request.status === 204 || request.status === 1223 || request.status === 201) {
            settings.success(request);
        }
        else {
            // Failure
            if (settings.error) {
                settings.error(errorHandler(request));
            }
            else {
                errorHandler(request);
            }
        }
    }
}