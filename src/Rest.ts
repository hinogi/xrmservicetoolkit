/// <reference path="../typings/main.d.ts" />
        // Create: createRecord,
        // Retrieve: retrieveRecord,
        // Update: updateRecord,
        // Delete: deleteRecord,
        // RetrieveMultiple: retrieveMultipleRecords,
        // Associate: associateRecord,
        // Disassociate: disassociateRecord
export default class Rest {


    var getXhr = function () {
        ///<summary>
        /// Get an instance of XMLHttpRequest for all browsers
        ///</summary>
        if (XMLHttpRequest) {
            // Chrome, Firefox, IE7+, Opera, Safari
            // ReSharper disable InconsistentNaming
            return new XMLHttpRequest();
            // ReSharper restore InconsistentNaming
        }
        // IE6
        try {
            // The latest stable version. It has the best security, performance,
            // reliability, and W3C conformance. Ships with Vista, and available
            // with other OS's via downloads and updates.
            return new ActiveXObject('MSXML2.XMLHTTP.6.0');
        } catch (e) {
            try {
                // The fallback.
                return new ActiveXObject('MSXML2.XMLHTTP.3.0');
            } catch (e) {
                alertMessage('This browser is not AJAX enabled.');
                return null;
            }
        }
    };

    var createRecord = function (object, type, successCallback, errorCallback, async) {
        ///<summary>
        /// Sends synchronous/asynchronous request to create a new record.
        ///</summary>
        ///<param name="object" type="Object">
        /// A JavaScript object with properties corresponding to the Schema name of
        /// entity attributes that are valid for create operations.
        ///</param>
        parameterCheck(object, "XrmServiceToolkit.REST.createRecord requires the object parameter.");
        ///<param name="type" type="string">
        /// A String representing the name of the entity
        ///</param>
        stringParameterCheck(type, "XrmServiceToolkit.REST.createRecord requires the type parameter is a string.");
        ///<param name="successCallback" type="Function">
        /// The function that will be passed through and be called by a successful response.
        /// This function can accept the returned record as a parameter.
        /// </param>
        callbackParameterCheck(successCallback, "XrmServiceToolkit.REST.createRecord requires the successCallback is a function.");
        ///<param name="errorCallback" type="Function">
        /// The function that will be passed through and be called by a failed response.
        /// This function must accept an Error object as a parameter.
        /// </param>
        callbackParameterCheck(errorCallback, "XrmServiceToolkit.REST.createRecord requires the errorCallback is a function.");
        ///<param name="async" type="Boolean">
        /// A Boolean representing if the method should run asynchronously or synchronously
        /// true means asynchronously. false means synchronously
        /// </param>
        booleanParameterCheck(async, "XrmServiceToolkit.REST.createRecord requires the async is a boolean.");

        var req = getXhr();
        req.open("POST", oDataPath() + type, async);
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");

        if (async) {
            req.onreadystatechange = function () {
                if (this.readyState === 4 /* complete */) {
                    req.onreadystatechange = null;
                    if (req.status === 201) {
                        successCallback(JSON.parse(req.responseText, dateReviver).d);
                    }
                    else {
                        errorCallback(errorHandler(req));
                    }
                }
            };
            req.send(JSON.stringify(object));
        } else {
            req.send(JSON.stringify(object));
            if (req.status === 201) {
                successCallback(JSON.parse(req.responseText, dateReviver).d);
            }
            else {
                errorCallback(errorHandler(req));
            }
        }

    };

    var retrieveRecord = function (id, type, select, expand, successCallback, errorCallback, async) {
        ///<summary>
        /// Sends synchronous/asynchronous request to retrieve a record.
        ///</summary>
        ///<param name="id" type="String">
        /// A String representing the GUID value for the record to retrieve.
        ///</param>
        stringParameterCheck(id, "XrmServiceToolkit.REST.retrieveRecord requires the id parameter is a string.");
        ///<param name="type" type="string">
        /// A String representing the name of the entity
        ///</param>
        stringParameterCheck(type, "XrmServiceToolkit.REST.retrieveRecord requires the type parameter is a string.");
        ///<param name="select" type="String">
        /// A String representing the $select OData System Query Option to control which
        /// attributes will be returned. This is a comma separated list of Attribute names that are valid for retrieve.
        /// If null all properties for the record will be returned
        ///</param>
        if (select != null)
            stringParameterCheck(select, "XrmServiceToolkit.REST.retrieveRecord requires the select parameter is a string.");
        ///<param name="expand" type="String">
        /// A String representing the $expand OData System Query Option value to control which
        /// related records are also returned. This is a comma separated list of of up to 6 entity relationship names
        /// If null no expanded related records will be returned.
        ///</param>
        if (expand != null)
            stringParameterCheck(expand, "XrmServiceToolkit.REST.retrieveRecord requires the expand parameter is a string.");
        ///<param name="successCallback" type="Function">
        /// The function that will be passed through and be called by a successful response.
        /// This function must accept the returned record as a parameter.
        /// </param>
        callbackParameterCheck(successCallback, "XrmServiceToolkit.REST.retrieveRecord requires the successCallback parameter is a function.");
        ///<param name="errorCallback" type="Function">
        /// The function that will be passed through and be called by a failed response.
        /// This function must accept an Error object as a parameter.
        /// </param>
        callbackParameterCheck(errorCallback, "XrmServiceToolkit.REST.retrieveRecord requires the errorCallback parameter is a function.");
        ///<param name="async" type="Boolean">
        /// A Boolean representing if the method should run asynchronously or synchronously
        /// true means asynchronously. false means synchronously
        /// </param>
        booleanParameterCheck(async, "XrmServiceToolkit.REST.retrieveRecord requires the async parameter is a boolean.");

        var systemQueryOptions = "";

        if (select != null || expand != null) {
            systemQueryOptions = "?";
            if (select != null) {
                var selectString = "$select=" + select;
                if (expand != null) {
                    selectString = selectString + "," + expand;
                }
                systemQueryOptions = systemQueryOptions + selectString;
            }
            if (expand != null) {
                systemQueryOptions = systemQueryOptions + "&$expand=" + expand;
            }
        }

        var req = getXhr();
        req.open("GET", oDataPath() + type + "(guid'" + id + "')" + systemQueryOptions, async);
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");

        if (async) {
            req.onreadystatechange = function () {
                if (req.readyState === 4 /* complete */) {
                    if (req.status === 200) {
                        successCallback(JSON.parse(req.responseText, dateReviver).d);
                    }
                    else {
                        errorCallback(errorHandler(req));
                    }
                }
            };
            req.send();
        } else {
            req.send();
            if (req.status === 200) {
                successCallback(JSON.parse(req.responseText, dateReviver).d);
            }
            else {
                errorCallback(errorHandler(req));
            }
        }

    };

    var updateRecord = function (id, object, type, successCallback, errorCallback, async) {
        ///<summary>
        /// Sends synchronous/asynchronous request to update a record.
        ///</summary>
        ///<param name="id" type="String">
        /// A String representing the GUID value for the record to update.
        ///</param>
        stringParameterCheck(id, "XrmServiceToolkit.REST.updateRecord requires the id parameter.");
        ///<param name="object" type="Object">
        /// A JavaScript object with properties corresponding to the Schema name of
        /// entity attributes that are valid for create operations.
        ///</param>
        parameterCheck(object, "XrmServiceToolkit.REST.updateRecord requires the object parameter.");
        ///<param name="type" type="string">
        /// A String representing the name of the entity
        ///</param>
        stringParameterCheck(type, "XrmServiceToolkit.REST.updateRecord requires the type parameter.");
        ///<param name="successCallback" type="Function">
        /// The function that will be passed through and be called by a successful response.
        /// Nothing will be returned to this function.
        /// </param>
        callbackParameterCheck(successCallback, "XrmServiceToolkit.REST.updateRecord requires the successCallback is a function.");
        ///<param name="errorCallback" type="Function">
        /// The function that will be passed through and be called by a failed response.
        /// This function must accept an Error object as a parameter.
        /// </param>
        callbackParameterCheck(errorCallback, "XrmServiceToolkit.REST.updateRecord requires the errorCallback is a function.");
        ///<param name="async" type="Boolean">
        /// A Boolean representing if the method should run asynchronously or synchronously
        /// true means asynchronously. false means synchronously
        /// </param>
        booleanParameterCheck(async, "XrmServiceToolkit.REST.updateRecord requires the async parameter is a boolean.");

        var req = getXhr();

        req.open("POST", oDataPath() + type + "(guid'" + id + "')", async);
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.setRequestHeader("X-HTTP-Method", "MERGE");

        if (async) {
            req.onreadystatechange = function () {
                if (req.readyState === 4 /* complete */) {
                    if (req.status === 204 || req.status === 1223) {
                        successCallback();
                    }
                    else {
                        errorCallback(errorHandler(req));
                    }
                }
            };
            req.send(JSON.stringify(object));
        } else {
            req.send(JSON.stringify(object));
            if (req.status === 204 || req.status === 1223) {
                successCallback();
            }
            else {
                errorCallback(errorHandler(req));
            }
        }

    };

    var deleteRecord = function (id, type, successCallback, errorCallback, async) {
        ///<summary>
        /// Sends synchronous/asynchronous request to delete a record.
        ///</summary>
        ///<param name="id" type="String">
        /// A String representing the GUID value for the record to delete.
        ///</param>
        stringParameterCheck(id, "XrmServiceToolkit.REST.deleteRecord requires the id parameter.");
        ///<param name="type" type="string">
        /// A String representing the name of the entity
        ///</param>
        stringParameterCheck(type, "XrmServiceToolkit.REST.deleteRecord requires the type parameter.");
        ///<param name="successCallback" type="Function">
        /// The function that will be passed through and be called by a successful response.
        /// Nothing will be returned to this function.
        /// </param>
        callbackParameterCheck(successCallback, "XrmServiceToolkit.REST.deleteRecord requires the successCallback is a function.");
        ///<param name="errorCallback" type="Function">
        /// The function that will be passed through and be called by a failed response.
        /// This function must accept an Error object as a parameter.
        /// </param>
        callbackParameterCheck(errorCallback, "XrmServiceToolkit.REST.deleteRecord requires the errorCallback is a function.");
        ///<param name="async" type="Boolean">
        /// A Boolean representing if the method should run asynchronously or synchronously
        /// true means asynchronously. false means synchronously
        /// </param>
        booleanParameterCheck(async, "XrmServiceToolkit.REST.deleteRecord requires the async parameter is a boolean.");

        var req = getXhr();
        req.open("POST", oDataPath() + type + "(guid'" + id + "')", async);
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.setRequestHeader("X-HTTP-Method", "DELETE");

        if (async) {
            req.onreadystatechange = function () {
                if (req.readyState === 4 /* complete */) {
                    if (req.status === 204 || req.status === 1223) {
                        successCallback();
                    }
                    else {
                        errorCallback(errorHandler(req));
                    }
                }
            };
            req.send();
        } else {
            req.send();
            if (req.status === 204 || req.status === 1223) {
                successCallback();
            }
            else {
                errorCallback(errorHandler(req));
            }
        }
    };

    var retrieveMultipleRecords = function (type, options, successCallback, errorCallback, onComplete, async) {
        ///<summary>
        /// Sends synchronous/asynchronous request to retrieve records.
        ///</summary>
        ///<param name="type" type="String">
        /// The Schema Name of the Entity type record to retrieve.
        /// For an Account record, use "Account"
        ///</param>
        stringParameterCheck(type, "XrmServiceToolkit.REST.retrieveMultipleRecords requires the type parameter is a string.");
        ///<param name="options" type="String">
        /// A String representing the OData System Query Options to control the data returned
        ///</param>
        if (options != null)
            stringParameterCheck(options, "XrmServiceToolkit.REST.retrieveMultipleRecords requires the options parameter is a string.");
        ///<param name="successCallback" type="Function">
        /// The function that will be passed through and be called for each page of records returned.
        /// Each page is 50 records. If you expect that more than one page of records will be returned,
        /// this function should loop through the results and push the records into an array outside of the function.
        /// Use the OnComplete event handler to know when all the records have been processed.
        /// </param>
        callbackParameterCheck(successCallback, "XrmServiceToolkit.REST.retrieveMultipleRecords requires the successCallback parameter is a function.");
        ///<param name="errorCallback" type="Function">
        /// The function that will be passed through and be called by a failed response.
        /// This function must accept an Error object as a parameter.
        /// </param>
        callbackParameterCheck(errorCallback, "XrmServiceToolkit.REST.retrieveMultipleRecords requires the errorCallback parameter is a function.");
        ///<param name="OnComplete" type="Function">
        /// The function that will be called when all the requested records have been returned.
        /// No parameters are passed to this function.
        /// </param>
        callbackParameterCheck(onComplete, "XrmServiceToolkit.REST.retrieveMultipleRecords requires the OnComplete parameter is a function.");
        ///<param name="async" type="Boolean">
        /// A Boolean representing if the method should run asynchronously or synchronously
        /// true means asynchronously. false means synchronously
        /// </param>
        booleanParameterCheck(async, "XrmServiceToolkit.REST.retrieveMultipleRecords requires the async parameter is a boolean.");

        var optionsString = '';
        if (options != null) {
            if (options.charAt(0) !== "?") {
                optionsString = "?" + options;
            }
            else {
                optionsString = options;
            }
        }

        var req = getXhr();
        req.open("GET", oDataPath() + type + optionsString, async);
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");

        if (async) {
            req.onreadystatechange = function () {
                if (req.readyState === 4 /* complete */) {
                    if (req.status === 200) {
                        var returned = JSON.parse(req.responseText, dateReviver).d;
                        successCallback(returned.results);
                        if (returned.__next == null) {
                            onComplete();
                        } else {
                            var queryOptions = returned.__next.substring((oDataPath() + type).length);
                            retrieveMultipleRecords(type, queryOptions, successCallback, errorCallback, onComplete, async);
                        }
                    }
                    else {
                        errorCallback(errorHandler(req));
                    }
                }
            };
            req.send();
        } else {
            req.send();
            if (req.status === 200) {
                var returned = JSON.parse(req.responseText, dateReviver).d;
                successCallback(returned.results);
                if (returned.__next == null) {
                    onComplete();
                } else {
                    var queryOptions = returned.__next.substring((oDataPath() + type).length);
                    retrieveMultipleRecords(type, queryOptions, successCallback, errorCallback, onComplete, async);
                }
            }
            else {
                errorCallback(errorHandler(req));
            }
        }
    };

    var performRequest = function (settings) {
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

    };

    var associateRecord = function (entityid1, odataSetName1, entityid2, odataSetName2, relationship, successCallback, errorCallback, async) {
        ///<summary>
        /// Sends synchronous/asynchronous request to associate a record.
        ///</summary>
        ///<param name="entityid1" type="string">
        /// A String representing the GUID value for the record to associate.
        ///</param>
        parameterCheck(entityid1, "XrmServiceToolkit.REST.associateRecord requires the entityid1 parameter.");
        ///<param name="odataSetName1" type="string">
        /// A String representing the odataset name for entityid1
        ///</param>
        parameterCheck(odataSetName1, "XrmServiceToolkit.REST.associateRecord requires the odataSetName1 parameter.");
        ///<param name="entityid2" type="string">
        /// A String representing the GUID value for the record to be associated.
        ///</param>
        parameterCheck(entityid2, "XrmServiceToolkit.REST.associateRecord requires the entityid2 parameter.");
        ///<param name="odataSetName2" type="string">
        /// A String representing the odataset name for entityid2
        ///</param>
        parameterCheck(odataSetName2, "XrmServiceToolkit.REST.associateRecord requires the odataSetName2 parameter.");
        ///<param name="relationship" type="string">
        /// A String representing the name of the relationship for association
        ///</param>
        parameterCheck(relationship, "XrmServiceToolkit.REST.associateRecord requires the relationship parameter.");
        ///<param name="successCallback" type="Function">
        /// The function that will be passed through and be called by a successful response.
        /// Nothing will be returned to this function.
        /// </param>
        callbackParameterCheck(successCallback, "XrmServiceToolkit.REST.associateRecord requires the successCallback is a function.");
        ///<param name="errorCallback" type="Function">
        /// The function that will be passed through and be called by a failed response.
        /// This function must accept an Error object as a parameter.
        /// </param>
        callbackParameterCheck(errorCallback, "XrmServiceToolkit.REST.associateRecord requires the errorCallback is a function.");
        ///<param name="async" type="Boolean">
        /// A Boolean representing if the method should run asynchronously or synchronously
        /// true means asynchronously. false means synchronously
        /// </param>
        booleanParameterCheck(async, "XrmServiceToolkit.REST.associateRecord requires the async parameter is a boolean");

        var entity2 = {};
        entity2.uri = oDataPath() + "/" + odataSetName2 + "(guid'" + entityid2 + "')";
        var jsonEntity = window.JSON.stringify(entity2);

        performRequest({
            type: "POST",
            url: oDataPath() + "/" + odataSetName1 + "(guid'" + entityid1 + "')/$links/" + relationship,
            data: jsonEntity,
            success: successCallback,
            error: errorCallback,
            async: async
        });
    };

    var disassociateRecord = function (entityid1, odataSetName, entityid2, relationship, successCallback, errorCallback, async) {
        ///<summary>
        /// Sends synchronous/asynchronous request to disassociate a record.
        ///</summary>
        ///<param name="entityid1" type="string">
        /// A String representing the GUID value for the record to disassociate.
        ///</param>
        parameterCheck(entityid1, "XrmServiceToolkit.REST.disassociateRecord requires the entityid1 parameter.");
        ///<param name="odataSetName" type="string">
        /// A String representing the odataset name for entityid1
        ///</param>
        parameterCheck(odataSetName, "XrmServiceToolkit.REST.disassociateRecord requires the odataSetName parameter.");
        ///<param name="entityid2" type="string">
        /// A String representing the GUID value for the record to be disassociated.
        ///</param>
        parameterCheck(entityid2, "XrmServiceToolkit.REST.disassociateRecord requires the entityid2 parameter.");
        ///<param name="relationship" type="string">
        /// A String representing the name of the relationship for disassociation
        ///</param>
        parameterCheck(relationship, "XrmServiceToolkit.REST.disassociateRecord requires the relationship parameter.");
        ///<param name="successCallback" type="Function">
        /// The function that will be passed through and be called by a successful response.
        /// Nothing will be returned to this function.
        /// </param>
        callbackParameterCheck(successCallback, "XrmServiceToolkit.REST.disassociateRecord requires the successCallback is a function.");
        ///<param name="errorCallback" type="Function">
        /// The function that will be passed through and be called by a failed response.
        /// This function must accept an Error object as a parameter.
        /// </param>
        callbackParameterCheck(errorCallback, "XrmServiceToolkit.REST.disassociateRecord requires the errorCallback is a function.");
        ///<param name="async" type="Boolean">
        /// A Boolean representing if the method should run asynchronously or synchronously
        /// true means asynchronously. false means synchronously
        /// </param>
        booleanParameterCheck(async, "XrmServiceToolkit.REST.disassociateRecord requires the async parameter is a boolean.");

        var url = oDataPath() + "/" + odataSetName + "(guid'" + entityid1 + "')/$links/" + relationship + "(guid'" + entityid2 + "')";
        performRequest({
            url: url,
            type: "POST",
            action: "DELETE",
            error: errorCallback,
            success: successCallback,
            async: async
        });
    };
}