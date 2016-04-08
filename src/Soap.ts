/// <reference path="../typings/main.d.ts" />
import {alertMessage, htmlEncode, innerSurrogateAmpersandWorkaround, crmXmlDecode, crmXmlEncode, encodeValue} from "./Helper";
import {xrmEntityReference, businessEntity, doRequest, xmlParser, xmlToString, selectSingleNodeText, selectSingleNode, fetchMore} from "./HelperSoap";

        // RetrieveMultiple: retrieveMultiple,
        // QueryByAttribute: queryByAttribute,
        // QueryAll: queryAll,
        // SetState: setState,
        // Assign: assign,
        // RetrievePrincipalAccess: retrievePrincipalAccess,
        // GrantAccess: grantAccess,
        // ModifyAccess: modifyAccess,
        // RevokeAccess: revokeAccess,
        // GetCurrentUserId: getCurrentUserId,
        // GetCurrentUserBusinessUnitId: getCurrentUserBusinessUnitId,
        // GetCurrentUserRoles: getCurrentUserRoles,
        // IsCurrentUserRole: isCurrentUserInRole,
        // RetrieveAllEntitiesMetadata: retrieveAllEntitiesMetadata,
        // RetrieveEntityMetadata: retrieveEntityMetadata,
        // RetrieveAttributeMetadata: retrieveAttributeMetadata
export default class Soap{

    /**
     * Sends synchronous/asynchronous request to create a new record
     *
     * @param {Object} be A JavaScript object with properties corresponding to the Schema name of
     * entity attributes that are valid for create operations.
     * @param {Function} callback A Function used for asynchronous request. If not defined, it sends a synchronous request
     * @returns {(void | any)} If sync -> results
     */
    static Create(be: businessEntity, callback?: Function): void | any {
        let request = be.serialize();
        let async = !!callback;
        let mBody = `
            <request i:type="a:CreateRequest" xmlns:a="http://schemas.microsoft.com/xrm/2011/Contracts">
                <a:Parameters xmlns:b="http://schemas.datacontract.org/2004/07/System.Collections.Generic">
                <a:KeyValuePairOfstringanyType>
                    <b:key>Target</b:key>
                    ${request}
                </a:KeyValuePairOfstringanyType>
                </a:Parameters>
                <a:RequestId i:nil="true" />
                <a:RequestName>Create</a:RequestName>
            </request>
        `;

        return doRequest(mBody, "Execute", async, (resultXml: string) => {
            let responseText = selectSingleNodeText(resultXml, "//b:value");
            let result = crmXmlDecode(responseText);

            if (!async) {
                return result;
            } else {
                callback(result);
            }
            // ReSharper disable NotAllPathsReturnValue
        });
        // ReSharper restore NotAllPathsReturnValue
    }

    /**
     * Sends synchronous/asynchronous request to update an existing record
     *
     * @param {businessEntity} be A JavaScript object with properties corresponding to the Schema name of
     * entity attributes that are valid for update operations
     * @param {Function} callback A Function used for asynchronous request. If not defined, it sends a synchronous request
     * @returns {(void | any)} If sync -> results
     */
    static Update(be: businessEntity, callback?: Function): void | any {
        let request = be.serialize();
        let async = !!callback;
        let mBody = `
            <request i:type="a:UpdateRequest" xmlns:a="http://schemas.microsoft.com/xrm/2011/Contracts">
                <a:Parameters xmlns:b="http://schemas.datacontract.org/2004/07/System.Collections.Generic">
                    <a:KeyValuePairOfstringanyType>
                        <b:key>Target</b:key>
                        ${request}
                    </a:KeyValuePairOfstringanyType>
                </a:Parameters>
                <a:RequestId i:nil="true" />
                <a:RequestName>Update</a:RequestName>
            </request>
        `;

        return doRequest(mBody, "Execute", async, (resultXml: string) => {
            let responseText = selectSingleNodeText(resultXml, "//a:Results");
            let result = crmXmlDecode(responseText);

            if (!async) {
                return result;
            } else {
                callback(result);
            }
            // ReSharper disable NotAllPathsReturnValue
        });
        // ReSharper restore NotAllPathsReturnValue
    }

    /**
     * Sends synchronous/asynchronous request to delete a record
     *
     * @param {string} entityName A JavaScript String corresponding to the Schema name of
     * entity that is used for delete operations
     * @param {string} id A JavaScript String corresponding to the GUID of
     * entity that is used for delete operations
     * @param {Function} [callback] A Function used for asynchronous request. If not defined, it sends a synchronous request
     * @returns {(void | any)} If sync -> results
     */
    static Delete(entityName: string, id: string, callback?: Function): void | any {
        let request =`
            <request i:type="a:DeleteRequest" xmlns:a="http://schemas.microsoft.com/xrm/2011/Contracts">
                <a:Parameters xmlns:b="http://schemas.datacontract.org/2004/07/System.Collections.Generic">
                    <a:KeyValuePairOfstringanyType>
                        <b:key>Target</b:key>
                        <b:value i:type="a:EntityReference">
                            <a:Id>"
                                ${id}
                            </a:Id>
                            <a:LogicalName>
                                ${entityName}
                            </a:LogicalName>
                            <a:Name i:nil="true" />
                        </b:value>
                    </a:KeyValuePairOfstringanyType>
                </a:Parameters>
                <a:RequestId i:nil="true" />
                <a:RequestName>Delete</a:RequestName>
            </request>
        `;
        let async = !!callback;

        return doRequest(request, "Execute", async, (resultXml: string) {
            let responseText = selectSingleNodeText(resultXml, "//a:Results");
            let result = crmXmlDecode(responseText);

            if (!async) {
                return result;
            } else {
                callback(result);
            }
            // ReSharper disable NotAllPathsReturnValue
        });
        // ReSharper restore NotAllPathsReturnValue
    }

    /**
     * Sends synchronous/asynchronous request to execute a soap request
     *
     * @param {string} request A JavaScript string corresponding to the soap request
     * that are valid for execute operations
     * @param {Function} [callback] A Function used for asynchronous request. If not defined, it sends a synchronous request
     * @returns {(void | any)} If sync -> results
     */
    static Execute(request: string, callback?: Function): void | any {
        var async = !!callback;

        return doRequest(request, "Execute", async, (resultXml: string) {
            if (!async){
                return resultXml;
            } else{
                callback(resultXml);
            }
            // ReSharper disable NotAllPathsReturnValue
        });
        // ReSharper restore NotAllPathsReturnValue
    }

    /**
     * Sends synchronous/asynchronous request to do a fetch request)
     *
     * @param {string} fetchCore A JavaScript String containing serialized XML using the FetchXML schema.
     * For efficiency, start with the "entity" node
     * @param {boolean} fetchAll Switch to enable paging
     * @param {Function} callback A Function used for asynchronous request. If not defined, it sends a synchronous request
     * @returns {(void | any)} If sync -> results
     */
    static Fetch (fetchCore: string, fetchAll: boolean, callback: Function): void | any {
        let fetchXml = fetchCore;

        if (fetchCore.slice(0, 7) === "<entity") {
            fetchXml =`
                <fetch mapping="logical">
                    ${fetchCore.replace(/\"/g, "'")}
                </fetch>
            `;
        } else {
            let isAggregate = (fetchCore.indexOf("aggregate=") !== -1);
            let isLimitedReturn = (fetchCore.indexOf("page='1'") !== -1 && fetchCore.indexOf("count='") !== -1);

            let distinctPos = fetchCore.indexOf("distinct=");
            let isDistinct = (distinctPos !== -1);
            let valQuotes = fetchCore.substring(distinctPos + 9, distinctPos + 10);
            let distinctValue = isDistinct
                ? fetchCore.substring(fetchCore.indexOf("distinct=") + 10, fetchCore.indexOf(valQuotes, fetchCore.indexOf("distinct=") + 10))
                : "false";
            let xmlDoc = xmlParser(fetchCore);
            let fetchEntity = selectSingleNode(xmlDoc, "//entity");
            if (fetchEntity === null) {
                throw new Error("XrmServiceToolkit.Fetch: No 'entity' node in the provided FetchXML.");
            }
            let fetchCoreDom = fetchEntity;
            try {
                fetchCore = xmlToString(fetchCoreDom).replace(/\"/g, "'");
            } catch (error) {
                if (fetchCoreDom !== undefined && fetchCoreDom.xml) {
                    fetchCore = fetchCoreDom.xml.replace(/\"/g, "'");
                } else {
                    throw new Error("XrmServiceToolkit.Fetch: This client does not provide the necessary XML features to continue.");
                }
            }

            if (!isAggregate && !isLimitedReturn) {
                fetchXml = `
                    <fetch mapping="logical" distinct="${(isDistinct ? distinctValue : "false")}"'" >
                        ${fetchCore}
                    </fetch>
                `;
            }
        }

        let request = `
            <request i:type="a:RetrieveMultipleRequest" xmlns:a="http://schemas.microsoft.com/xrm/2011/Contracts">
                <a:Parameters xmlns:b="http://schemas.datacontract.org/2004/07/System.Collections.Generic">
                    <a:KeyValuePairOfstringanyType>
                        <b:key>Query</b:key>
                        <b:value i:type="a:FetchExpression">
                            <a:Query>${crmXmlEncode(fetchXml)}</a:Query>
                        </b:value>
                    </a:KeyValuePairOfstringanyType>
                </a:Parameters>
                <a:RequestId i:nil="true"/>
                <a:RequestName>RetrieveMultiple</a:RequestName>
            </request>
        `;

        let async = !!callback;

        return doRequest(request, "Execute", async, (resultXml: string) => {
            let fetchResult: Node = selectSingleNode(resultXml, "//a:Entities");
            let moreRecords: boolean = (selectSingleNodeText(resultXml, "//a:MoreRecords") === "true");

            let fetchResults: Array<any> = [];
            if (fetchResult != null) {
                for (let ii: number = 0, olength = fetchResult.childNodes.length; ii < olength; ii++) {
                    let entity: businessEntity = new businessEntity();

                    entity.deserialize(fetchResult.childNodes[ii]);
                    fetchResults.push(entity);
                }

                if (fetchAll && moreRecords) {
                    let pageCookie = selectSingleNodeText(resultXml, "//a:PagingCookie").replace(/\"/g, '\'').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/'/g, '&quot;');

                    fetchMore(fetchCore, 2, pageCookie, fetchResults);
                }

                if (!async){
                    return fetchResults;
                } else{
                    callback(fetchResults);
                }
            }
            // ReSharper disable once NotAllPathsReturnValue
        });
    }

    /**
     * Sends synchronous/asynchronous request to retrieve a record
     *
     * @param {string} entityName A JavaScript String corresponding to the Schema name of
     * entity that is used for retrieve operations
     * @param {string} id A JavaScript String corresponding to the GUID of
     * entity that is used for retrieve operations
     * @param {Array} columnSet  A JavaScript Array corresponding to the attributes of
     * entity that is used for retrieve operations
     * @param {Function} callback A Function used for asynchronous request. If not defined, it sends a synchronous request
     * @returns {(void | any)} If sync -> results
     */
    static Retrieve(entityName: string, id: string, columnSet: Array<any>, callback: Function): void | any {
        let attributes = "";
        // ReSharper disable AssignedValueIsNeverUsed
        let query = "";
        // ReSharper restore AssignedValueIsNeverUsed
        if (columnSet != null) {
            for (let i = 0, ilength = columnSet.length; i < ilength; i++) {
                attributes += "<c:string>" + columnSet[i] + "</c:string>";
            }
            query = "<a:AllColumns>false</a:AllColumns>" +
                    "<a:Columns xmlns:c='http://schemas.microsoft.com/2003/10/Serialization/Arrays'>" +
                        attributes +
                    "</a:Columns>";
        }
        else {
            query = "<a:AllColumns>true</a:AllColumns><a:Columns xmlns:b='http://schemas.microsoft.com/2003/10/Serialization/Arrays' />";
        }

        let msgBody = `
            <request i:type="a:RetrieveRequest" xmlns:a="http://schemas.microsoft.com/xrm/2011/Contracts">
                <a:Parameters xmlns:b="http://schemas.datacontract.org/2004/07/System.Collections.Generic">
                    <a:KeyValuePairOfstringanyType>
                        <b:key>Target</b:key>
                        <b:value i:type="a:EntityReference">
                            <a:Id>${encodeValue(id)}</a:Id>
                            <a:LogicalName>${entityName}</a:LogicalName>
                            <a:Name i:nil="true" />
                        </b:value>
                    </a:KeyValuePairOfstringanyType>
                    <a:KeyValuePairOfstringanyType>
                        <b:key>ColumnSet</b:key>
                        <b:value i:type="a:ColumnSet">
                            ${query}
                        </b:value>
                    </a:KeyValuePairOfstringanyType>
                </a:Parameters>
                <a:RequestId i:nil="true" />
                <a:RequestName>Retrieve</a:RequestName>
            </request>
        `;

        let async = !!callback;

        return doRequest(msgBody, "Execute", !!callback, (resultXml: string) => {
            let retrieveResult: Node = selectSingleNode(resultXml, "//b:value");
            let entity: businessEntity = new businessEntity();
            entity.deserialize(retrieveResult);

            if (!async){
                return entity;
            } else {
                callback(entity);
            }
            // ReSharper disable NotAllPathsReturnValue
        });
        // ReSharper restore NotAllPathsReturnValue
    }

    /**
     * Sends synchronous/asynchronous request to do a retrieveMultiple request
     *
     * @param {string} query A JavaScript String with properties corresponding to the retrievemultiple request
     * that are valid for retrievemultiple operations
     * @param {Function} callback A Function used for asynchronous request. If not defined, it sends a synchronous request
     * @returns {(void | any)} If sync -> results
     */
    static RetrieveMultiple(query: string, callback: Function): void | any {
        var request = `
            <request i:type="a:RetrieveMultipleRequest" xmlns:a="http://schemas.microsoft.com/xrm/2011/Contracts">
                <a:Parameters xmlns:b="http://schemas.datacontract.org/2004/07/System.Collections.Generic">
                    <a:KeyValuePairOfstringanyType>
                        <b:key>Query</b:key>
                        <b:value i:type="a:QueryExpression">
                            ${query}
                        </b:value>
                    </a:KeyValuePairOfstringanyType>
                </a:Parameters>
                <a:RequestId i:nil="true"/>
                <a:RequestName>RetrieveMultiple</a:RequestName>
            </request>
        `;

        let async = !!callback;

        return doRequest(request, "Execute", async, (resultXml: string) => {
            let resultNodes: Node = selectSingleNode(resultXml, "//a:Entities");

            let retriveMultipleResults: Array<businessEntity> = [];

            for (let i = 0, ilength = resultNodes.childNodes.length; i < ilength; i++) {
                let entity = new businessEntity();

                entity.deserialize(resultNodes.childNodes[i]);
                retriveMultipleResults[i] = entity;
            }

            if (!async){
                return retriveMultipleResults;
            } else{
                callback(retriveMultipleResults);
            }
            // ReSharper disable NotAllPathsReturnValue
        });
        // ReSharper restore NotAllPathsReturnValue
    }

    let joinArray = function (prefix, array, suffix) {
        let output = [];
        for (let i = 0, ilength = array.length; i < ilength; i++) {
            if (array[i] !== "" && array[i] != undefined) {
                output.push(prefix, array[i], suffix);
            }
        }
        return output.join("");
    };

    var joinConditionPair = function (attributes, values) {
        var output = [];
        for (var i = 0, ilength = attributes.length; i < ilength; i++) {
            if (attributes[i] !== "") {
                var value1 = values[i];
                if (typeof value1 == typeof []) {
                    output.push("<condition attribute='", attributes[i], "' operator='in' >");

                    for (var valueIndex in value1) {
                        if (value1.hasOwnProperty(valueIndex)) {
                            var value = encodeValue(value1[valueIndex]);
                            output.push("<value>" + value + "</value>");
                        }
                    }

                    output.push("</condition>");
                }
                else if (typeof value1 == typeof "") {
                    output.push("<condition attribute='", attributes[i], "' operator='eq' value='", encodeValue(value1), "' />");
                }
            }
        }
        return output.join("");
    };



    var queryByAttribute = function (queryOptions, callback) {
        ///<summary>
        /// Sends synchronous/asynchronous request to do a queryByAttribute request.
        ///</summary>
        ///<param name="queryOptions" type="Object">
        /// A JavaScript Object with properties corresponding to the queryByAttribute Criteria
        /// that are valid for queryByAttribute operations.
        /// queryOptions.entityName is a string represents the name of the entity
        /// queryOptions.attributes is a array represents the attributes of the entity to query
        /// queryOptions.values is a array represents the values of the attributes to query
        /// queryOptions.columnSet is a array represents the attributes of the entity to return
        /// queryOptions.orderBy is a array represents the order conditions of the results
        /// </param>
        ///<param name="callback" type="Function">
        /// A Function used for asynchronous request. If not defined, it sends a synchronous request.
        /// </param>
        var entityName = queryOptions.entityName;
        var attributes = queryOptions.attributes;
        var values = queryOptions.values;
        var columnSet = queryOptions.columnSet;
        var orderBy = queryOptions.orderBy || '';

        attributes = isArray(attributes) ? attributes : [attributes];
        values = isArray(values) ? values : [values];
        orderBy = (!!orderBy && isArray(orderBy)) ? orderBy : [orderBy];
        columnSet = (!!columnSet && isArray(columnSet)) ? columnSet : [columnSet];

        var xml =
                [
                    "<entity name='", entityName, "'>",
                            joinArray("<attribute name='", columnSet, "' />"),
                            joinArray("<order attribute='", orderBy, "' />"),
                        "<filter>",
                                joinConditionPair(attributes, values),
                        "</filter>",
                    "</entity>"
                ].join("");

        return fetch(xml, false, callback);
    };

    var queryAll = function (queryOptions, callback) {
        ///<summary>
        /// Sends synchronous/asynchronous request to do a queryAll request. This is to return all records (>5k+).
        /// Consider Performance impact when using this method.
        ///</summary>
        ///<param name="queryOptions" type="Object">
        /// A JavaScript Object with properties corresponding to the queryByAttribute Criteria
        /// that are valid for queryByAttribute operations.
        /// queryOptions.entityName is a string represents the name of the entity
        /// queryOptions.attributes is a array represents the attributes of the entity to query
        /// queryOptions.values is a array represents the values of the attributes to query
        /// queryOptions.columnSet is a array represents the attributes of the entity to return
        /// queryOptions.orderBy is a array represents the order conditions of the results
        /// </param>
        ///<param name="callback" type="Function">
        /// A Function used for asynchronous request. If not defined, it sends a synchronous request.
        /// </param>
        var entityName = queryOptions.entityName;
        var attributes = queryOptions.attributes;
        var values = queryOptions.values;
        var columnSet = queryOptions.columnSet;
        var orderBy = queryOptions.orderBy || '';

        attributes = isArray(attributes) ? attributes : [attributes];
        values = isArray(values) ? values : [values];
        orderBy = (!!orderBy && isArray(orderBy)) ? orderBy : [orderBy];
        columnSet = (!!columnSet && isArray(columnSet)) ? columnSet : [columnSet];

        var fetchCore = [
                    "<entity name='", entityName, "'>",
                            joinArray("<attribute name='", columnSet, "' />"),
                            joinArray("<order attribute='", orderBy, "' />"),
                        "<filter>",
                                joinConditionPair(attributes, values),
                        "</filter>",
                    "</entity>"
        ].join("");


        var async = !!callback;

        return fetch(fetchCore, true, async);
    };

    var setState = function (entityName, id, stateCode, statusCode, callback) {
        ///<summary>
        /// Sends synchronous/asynchronous request to setState of a record.
        ///</summary>
        ///<param name="entityName" type="String">
        /// A JavaScript String corresponding to the Schema name of
        /// entity that is used for setState operations.
        /// </param>
        ///<param name="id" type="String">
        /// A JavaScript String corresponding to the GUID of
        /// entity that is used for setState operations.
        /// </param>
        ///<param name="stateCode" type="Int">
        /// A JavaScript Integer corresponding to the value of
        /// entity state that is used for setState operations.
        /// </param>
        ///<param name="statusCode" type="Int">
        /// A JavaScript Integer corresponding to the value of
        /// entity status that is used for setState operations.
        /// </param>
        ///<param name="callback" type="Function">
        /// A Function used for asynchronous request. If not defined, it sends a synchronous request.
        /// </param>
        var request = [
            "<request i:type='b:SetStateRequest' xmlns:a='http://schemas.microsoft.com/xrm/2011/Contracts' xmlns:b='http://schemas.microsoft.com/crm/2011/Contracts'>",
                "<a:Parameters xmlns:c='http://schemas.datacontract.org/2004/07/System.Collections.Generic'>",
                    "<a:KeyValuePairOfstringanyType>",
                        "<c:key>EntityMoniker</c:key>",
                        "<c:value i:type='a:EntityReference'>",
                            "<a:Id>", encodeValue(id), "</a:Id>",
                            "<a:LogicalName>", entityName, "</a:LogicalName>",
                            "<a:Name i:nil='true' />",
                        "</c:value>",
                        "</a:KeyValuePairOfstringanyType>",
                        "<a:KeyValuePairOfstringanyType>",
                        "<c:key>State</c:key>",
                        "<c:value i:type='a:OptionSetValue'>",
                            "<a:Value>", stateCode.toString(), "</a:Value>",
                        "</c:value>",
                        "</a:KeyValuePairOfstringanyType>",
                        "<a:KeyValuePairOfstringanyType>",
                        "<c:key>Status</c:key>",
                        "<c:value i:type='a:OptionSetValue'>",
                            "<a:Value>", statusCode.toString(), "</a:Value>",
                        "</c:value>",
                        "</a:KeyValuePairOfstringanyType>",
                "</a:Parameters>",
                "<a:RequestId i:nil='true' />",
                "<a:RequestName>SetState</a:RequestName>",
            "</request>"
        ].join("");

        var async = !!callback;

        return doRequest(request, "Execute", async, function (resultXml) {
            var responseText = selectSingleNodeText(resultXml, "//ser:ExecuteResult");
            var result = crmXmlDecode(responseText);
            if (!async)
                return result;
            else
                callback(result);
            // ReSharper disable NotAllPathsReturnValue
        });
        // ReSharper restore NotAllPathsReturnValue
    };

    var associate = function (relationshipName, targetEntityName, targetId, relatedEntityName, relatedBusinessEntities, callback) {
        ///<summary>
        /// Sends synchronous/asynchronous request to associate records.
        ///</summary>
        ///<param name="relationshipName" type="String">
        /// A JavaScript String corresponding to the relationship name
        /// that is used for associate operations.
        /// </param>
        ///<param name="targetEntityName" type="String">
        /// A JavaScript String corresponding to the schema name of the target entity
        /// that is used for associate operations.
        /// </param>
        ///<param name="targetId" type="String">
        /// A JavaScript String corresponding to the GUID of the target entity
        /// that is used for associate operations.
        /// </param>
        ///<param name="relatedEntityName" type="String">
        /// A JavaScript String corresponding to the schema name of the related entity
        /// that is used for associate operations.
        /// </param>
        ///<param name="relationshipBusinessEntities" type="Array">
        /// A JavaScript Array corresponding to the collection of the related entities as BusinessEntity
        /// that is used for associate operations.
        /// </param>
        ///<param name="callback" type="Function">
        /// A Function used for asynchronous request. If not defined, it sends a synchronous request.
        /// </param>
        var relatedEntities = relatedBusinessEntities;

        relatedEntities = isArray(relatedEntities) ? relatedEntities : [relatedEntities];

        var output = [];
        for (var i = 0, ilength = relatedEntities.length; i < ilength; i++) {
            if (relatedEntities[i].id !== "") {
                output.push("<a:EntityReference>",
                                "<a:Id>", relatedEntities[i].id, "</a:Id>",
                                "<a:LogicalName>", relatedEntityName, "</a:LogicalName>",
                                "<a:Name i:nil='true' />",
                            "</a:EntityReference>");
            }
        }

        var relatedXml = output.join("");

        var request = [
            "<request i:type='a:AssociateRequest' xmlns:a='http://schemas.microsoft.com/xrm/2011/Contracts'>",
                "<a:Parameters xmlns:b='http://schemas.datacontract.org/2004/07/System.Collections.Generic'>",
                    "<a:KeyValuePairOfstringanyType>",
                        "<b:key>Target</b:key>",
                        "<b:value i:type='a:EntityReference'>",
                            "<a:Id>", encodeValue(targetId), "</a:Id>",
                            "<a:LogicalName>", targetEntityName, "</a:LogicalName>",
                            "<a:Name i:nil='true' />",
                        "</b:value>",
                    "</a:KeyValuePairOfstringanyType>",
                    "<a:KeyValuePairOfstringanyType>",
                        "<b:key>Relationship</b:key>",
                        "<b:value i:type='a:Relationship'>",
                            "<a:PrimaryEntityRole>Referenced</a:PrimaryEntityRole>",
                            "<a:SchemaName>", relationshipName, "</a:SchemaName>",
                        "</b:value>",
                    "</a:KeyValuePairOfstringanyType>",
                    "<a:KeyValuePairOfstringanyType>",
                    "<b:key>RelatedEntities</b:key>",
                    "<b:value i:type='a:EntityReferenceCollection'>",
                        relatedXml,
                    "</b:value>",
                    "</a:KeyValuePairOfstringanyType>",
                "</a:Parameters>",
                "<a:RequestId i:nil='true' />",
                "<a:RequestName>Associate</a:RequestName>",
            "</request>"
        ].join("");

        var async = !!callback;

        return doRequest(request, "Execute", async, function (resultXml) {
            var responseText = selectSingleNodeText(resultXml, "//ser:ExecuteResult");
            var result = crmXmlDecode(responseText);
            if (!async)
                return result;
            else
                callback(result);
            // ReSharper disable NotAllPathsReturnValue
        });
        // ReSharper restore NotAllPathsReturnValue
    };

    var disassociate = function (relationshipName, targetEntityName, targetId, relatedEntityName, relatedBusinessEntities, callback) {
        ///<summary>
        /// Sends synchronous/asynchronous request to disassociate records.
        ///</summary>
        ///<param name="relationshipName" type="String">
        /// A JavaScript String corresponding to the relationship name
        /// that is used for disassociate operations.
        /// </param>
        ///<param name="targetEntityName" type="String">
        /// A JavaScript String corresponding to the schema name of the target entity
        /// that is used for disassociate operations.
        /// </param>
        ///<param name="targetId" type="String">
        /// A JavaScript String corresponding to the GUID of the target entity
        /// that is used for disassociate operations.
        /// </param>
        ///<param name="relatedEntityName" type="String">
        /// A JavaScript String corresponding to the schema name of the related entity
        /// that is used for disassociate operations.
        /// </param>
        ///<param name="relationshipBusinessEntities" type="Array">
        /// A JavaScript Array corresponding to the collection of the related entities as BusinessEntity
        /// that is used for disassociate operations.
        /// </param>
        ///<param name="callback" type="Function">
        /// A Function used for asynchronous request. If not defined, it sends a synchronous request.
        /// </param>
        var relatedEntities = relatedBusinessEntities;

        relatedEntities = isArray(relatedEntities) ? relatedEntities : [relatedEntities];

        var output = [];
        for (var i = 0, ilength = relatedEntities.length; i < ilength; i++) {
            if (relatedEntities[i].id !== "") {
                output.push("<a:EntityReference>",
                                "<a:Id>", relatedEntities[i].id, "</a:Id>",
                                "<a:LogicalName>", relatedEntityName, "</a:LogicalName>",
                                "<a:Name i:nil='true' />",
                            "</a:EntityReference>");
            }
        }

        var relatedXml = output.join("");

        var request = [
            "<request i:type='a:DisassociateRequest' xmlns:a='http://schemas.microsoft.com/xrm/2011/Contracts'>",
                "<a:Parameters xmlns:b='http://schemas.datacontract.org/2004/07/System.Collections.Generic'>",
                    "<a:KeyValuePairOfstringanyType>",
                        "<b:key>Target</b:key>",
                        "<b:value i:type='a:EntityReference'>",
                            "<a:Id>", encodeValue(targetId), "</a:Id>",
                            "<a:LogicalName>", targetEntityName, "</a:LogicalName>",
                            "<a:Name i:nil='true' />",
                        "</b:value>",
                    "</a:KeyValuePairOfstringanyType>",
                    "<a:KeyValuePairOfstringanyType>",
                        "<b:key>Relationship</b:key>",
                        "<b:value i:type='a:Relationship'>",
                            "<a:PrimaryEntityRole i:nil='true' />",
                            "<a:SchemaName>", relationshipName, "</a:SchemaName>",
                        "</b:value>",
                    "</a:KeyValuePairOfstringanyType>",
                    "<a:KeyValuePairOfstringanyType>",
                    "<b:key>RelatedEntities</b:key>",
                    "<b:value i:type='a:EntityReferenceCollection'>",
                        relatedXml,
                    "</b:value>",
                    "</a:KeyValuePairOfstringanyType>",
                "</a:Parameters>",
                "<a:RequestId i:nil='true' />",
                "<a:RequestName>Disassociate</a:RequestName>",
            "</request>"
        ].join("");

        var async = !!callback;

        return doRequest(request, "Execute", async, function (resultXml) {
            var responseText = selectSingleNodeText(resultXml, "//ser:ExecuteResult");
            var result = crmXmlDecode(responseText);
            if (!async)
                return result;
            else
                callback(result);
            // ReSharper disable NotAllPathsReturnValue
        });
        // ReSharper restore NotAllPathsReturnValue
    };

    var getCurrentUserId = function () {
        ///<summary>
        /// Sends synchronous request to retrieve the GUID of the current user.
        ///</summary>
        var request = [
                "<request i:type='b:WhoAmIRequest' xmlns:a='http://schemas.microsoft.com/xrm/2011/Contracts' xmlns:b='http://schemas.microsoft.com/crm/2011/Contracts'>",
                "<a:Parameters xmlns:c='http://schemas.datacontract.org/2004/07/System.Collections.Generic' />",
                "<a:RequestId i:nil='true' />",
                "<a:RequestName>WhoAmI</a:RequestName>",
                "</request>"
        ].join("");
        var xmlDoc = doRequest(request, "Execute");

        return getNodeText(selectNodes(xmlDoc, "//b:value")[0]);

    };

    var getCurrentUserBusinessUnitId = function () {
        ///<summary>
        /// Sends synchronous request to retrieve the GUID of the current user's business unit.
        ///</summary>
        var request = ["<request i:type='b:WhoAmIRequest' xmlns:a='http://schemas.microsoft.com/xrm/2011/Contracts' xmlns:b='http://schemas.microsoft.com/crm/2011/Contracts'>",
                        "<a:Parameters xmlns:c='http://schemas.datacontract.org/2004/07/System.Collections.Generic' />",
                        "<a:RequestId i:nil='true' />",
                        "<a:RequestName>WhoAmI</a:RequestName>",
                        "</request>"].join("");
        var xmlDoc = doRequest(request, "Execute");

        return getNodeText(selectNodes(xmlDoc, "//b:value")[1]);
    };

    var getCurrentUserRoles = function () {
        ///<summary>
        /// Sends synchronous request to retrieve the list of the current user's roles.
        ///</summary>
        var xml =
                [
                    "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>",
                        "<entity name='role'>",
                        "<attribute name='name' />",
                        "<attribute name='businessunitid' />",
                        "<attribute name='roleid' />",
                        "<order attribute='name' descending='false' />" +
                        "<link-entity name='systemuserroles' from='roleid' to='roleid' visible='false' intersect='true'>",
                            "<link-entity name='systemuser' from='systemuserid' to='systemuserid' alias='aa'>",
                            "<filter type='and'>",
                                "<condition attribute='systemuserid' operator='eq-userid' />",
                            "</filter>",
                            "</link-entity>",
                        "</link-entity>",
                        "</entity>",
                    "</fetch>"
                ].join("");

        var fetchResult = fetch(xml);
        var roles = [];

        if (fetchResult !== null && typeof fetchResult != 'undefined') {
            for (var i = 0, ilength = fetchResult.length; i < ilength; i++) {
                roles[i] = fetchResult[i].attributes["name"].value;
            }
        }

        return roles;
    };

    var isCurrentUserInRole = function () {
        ///<summary>
        /// Sends synchronous request to check if the current user has certain roles
        /// Passes name of role as arguments. For example, IsCurrentUserInRole('System Administrator')
        /// Returns true or false.
        ///</summary>
        var roles = getCurrentUserRoles();
        for (var i = 0, ilength = roles.length; i < ilength; i++) {
            for (var j = 0, jlength = arguments.length; j < jlength; j++) {
                if (roles[i] === arguments[j]) {
                    return true;
                }
            }
        }

        return false;
    };

    var assign = function (targetEntityName, targetId, assigneeEntityName, assigneeId, callback) {
        ///<summary>
        /// Sends synchronous/asynchronous request to assign an existing record to a user / a team.
        ///</summary>
        ///<param name="targetEntityName" type="String">
        /// A JavaScript String corresponding to the schema name of the target entity
        /// that is used for assign operations.
        /// </param>
        ///<param name="targetId" type="String">
        /// A JavaScript String corresponding to the GUID of the target entity
        /// that is used for assign operations.
        /// </param>
        ///<param name="assigneeEntityName" type="String">
        /// A JavaScript String corresponding to the schema name of the assignee entity
        /// that is used for assign operations.
        /// </param>
        ///<param name="assigneeId" type="String">
        /// A JavaScript String corresponding to the GUID of the assignee entity
        /// that is used for assign operations.
        /// </param>
        ///<param name="callback" type="Function">
        /// A Function used for asynchronous request. If not defined, it sends a synchronous request.
        /// </param>

        var request = ["<request i:type='b:AssignRequest' xmlns:a='http://schemas.microsoft.com/xrm/2011/Contracts' xmlns:b='http://schemas.microsoft.com/crm/2011/Contracts'>",
                        "<a:Parameters xmlns:c='http://schemas.datacontract.org/2004/07/System.Collections.Generic'>",
                            "<a:KeyValuePairOfstringanyType>",
                            "<c:key>Target</c:key>",
                            "<c:value i:type='a:EntityReference'>",
                                "<a:Id>", encodeValue(targetId), "</a:Id>",
                                "<a:LogicalName>", targetEntityName, "</a:LogicalName>",
                                "<a:Name i:nil='true' />",
                            "</c:value>",
                            "</a:KeyValuePairOfstringanyType>",
                            "<a:KeyValuePairOfstringanyType>",
                            "<c:key>Assignee</c:key>",
                            "<c:value i:type='a:EntityReference'>",
                                "<a:Id>", encodeValue(assigneeId), "</a:Id>",
                                "<a:LogicalName>", assigneeEntityName, "</a:LogicalName>",
                                "<a:Name i:nil='true' />",
                            "</c:value>",
                            "</a:KeyValuePairOfstringanyType>",
                        "</a:Parameters>",
                        "<a:RequestId i:nil='true' />",
                        "<a:RequestName>Assign</a:RequestName>",
                        "</request>"].join("");
        var async = !!callback;

        return doRequest(request, "Execute", async, function (resultXml) {
            var responseText = selectSingleNodeText(resultXml, "//ser:ExecuteResult");
            var result = crmXmlDecode(responseText);
            if (!async)
                return result;
            else
                callback(result);
            // ReSharper disable NotAllPathsReturnValue
        });
        // ReSharper restore NotAllPathsReturnValue
    };

    var grantAccess = function (accessOptions, callback) {
        ///<summary>
        /// Sends synchronous/asynchronous request to do a grantAccess request.
        /// Levels of Access Options are: AppendAccess, AppendToAccess, AssignAccess, CreateAccess, DeleteAccess, None, ReadAccess, ShareAccess, and WriteAccess
        ///</summary>
        ///<param name="accessOptions" type="Object">
        /// A JavaScript Object with properties corresponding to the grantAccess Criteria
        /// that are valid for grantAccess operations.
        /// accessOptions.targetEntityName is a string represents the name of the target entity
        /// accessOptions.targetEntityId is a string represents the GUID of the target entity
        /// accessOptions.principalEntityName is a string represents the name of the principal entity
        /// accessOptions.principalEntityId is a string represents the GUID of the principal entity
        /// accessOptions.accessRights is a array represents the access conditions of the results
        /// </param>
        ///<param name="callback" type="Function">
        /// A Function used for asynchronous request. If not defined, it sends a synchronous request.
        /// </param>

        var targetEntityName = accessOptions.targetEntityName;
        var targetEntityId = accessOptions.targetEntityId;
        var principalEntityName = accessOptions.principalEntityName;
        var principalEntityId = accessOptions.principalEntityId;
        var accessRights = accessOptions.accessRights;

        accessRights = isArray(accessRights) ? accessRights : [accessRights];

        var accessRightString = "";
        for (var i = 0, ilength = accessRights.length; i < ilength; i++) {
            accessRightString += encodeValue(accessRights[i]) + " ";
        }

        var request = ["<request i:type='b:GrantAccessRequest' xmlns:a='http://schemas.microsoft.com/xrm/2011/Contracts' xmlns:b='http://schemas.microsoft.com/crm/2011/Contracts'>",
                        "<a:Parameters xmlns:c='http://schemas.datacontract.org/2004/07/System.Collections.Generic'>",
                            "<a:KeyValuePairOfstringanyType>",
                            "<c:key>Target</c:key>",
                            "<c:value i:type='a:EntityReference'>",
                                "<a:Id>", encodeValue(targetEntityId), "</a:Id>",
                                "<a:LogicalName>", targetEntityName, "</a:LogicalName>",
                                "<a:Name i:nil='true' />",
                            "</c:value>",
                            "</a:KeyValuePairOfstringanyType>",
                            "<a:KeyValuePairOfstringanyType>",
                            "<c:key>PrincipalAccess</c:key>",
                            "<c:value i:type='b:PrincipalAccess'>",
                                "<b:AccessMask>", accessRightString, "</b:AccessMask>",
                                "<b:Principal>",
                                "<a:Id>", encodeValue(principalEntityId), "</a:Id>",
                                "<a:LogicalName>", principalEntityName, "</a:LogicalName>",
                                "<a:Name i:nil='true' />",
                                "</b:Principal>",
                            "</c:value>",
                            "</a:KeyValuePairOfstringanyType>",
                        "</a:Parameters>",
                        "<a:RequestId i:nil='true' />",
                        "<a:RequestName>GrantAccess</a:RequestName>",
                    "</request>"].join("");
        var async = !!callback;

        return doRequest(request, "Execute", async, function (resultXml) {
            var responseText = selectSingleNodeText(resultXml, "//ser:ExecuteResult");
            var result = crmXmlDecode(responseText);
            if (!async)
                return result;
            else
                callback(result);
            // ReSharper disable NotAllPathsReturnValue
        });
        // ReSharper restore NotAllPathsReturnValue
    };

    var modifyAccess = function (accessOptions, callback) {
        ///<summary>
        /// Sends synchronous/asynchronous request to do a modifyAccess request.
        /// Levels of Access Options are: AppendAccess, AppendToAccess, AssignAccess, CreateAccess, DeleteAccess, None, ReadAccess, ShareAccess, and WriteAccess
        ///</summary>
        ///<param name="accessOptions" type="Object">
        /// A JavaScript Object with properties corresponding to the modifyAccess Criteria
        /// that are valid for modifyAccess operations.
        /// accessOptions.targetEntityName is a string represents the name of the target entity
        /// accessOptions.targetEntityId is a string represents the GUID of the target entity
        /// accessOptions.principalEntityName is a string represents the name of the principal entity
        /// accessOptions.principalEntityId is a string represents the GUID of the principal entity
        /// accessOptions.accessRights is a array represents the access conditions of the results
        /// </param>
        ///<param name="callback" type="Function">
        /// A Function used for asynchronous request. If not defined, it sends a synchronous request.
        /// </param>

        var targetEntityName = accessOptions.targetEntityName;
        var targetEntityId = accessOptions.targetEntityId;
        var principalEntityName = accessOptions.principalEntityName;
        var principalEntityId = accessOptions.principalEntityId;
        var accessRights = accessOptions.accessRights;

        accessRights = isArray(accessRights) ? accessRights : [accessRights];

        var accessRightString = "";
        for (var i = 0, ilength = accessRights.length; i < ilength; i++) {
            accessRightString += encodeValue(accessRights[i]) + " ";
        }

        var request = ["<request i:type='b:ModifyAccessRequest' xmlns:a='http://schemas.microsoft.com/xrm/2011/Contracts' xmlns:b='http://schemas.microsoft.com/crm/2011/Contracts'>",
                        "<a:Parameters xmlns:c='http://schemas.datacontract.org/2004/07/System.Collections.Generic'>",
                            "<a:KeyValuePairOfstringanyType>",
                            "<c:key>Target</c:key>",
                            "<c:value i:type='a:EntityReference'>",
                                "<a:Id>", encodeValue(targetEntityId), "</a:Id>",
                                "<a:LogicalName>", targetEntityName, "</a:LogicalName>",
                                "<a:Name i:nil='true' />",
                            "</c:value>",
                            "</a:KeyValuePairOfstringanyType>",
                            "<a:KeyValuePairOfstringanyType>",
                            "<c:key>PrincipalAccess</c:key>",
                            "<c:value i:type='b:PrincipalAccess'>",
                                "<b:AccessMask>", accessRightString, "</b:AccessMask>",
                                "<b:Principal>",
                                "<a:Id>", encodeValue(principalEntityId), "</a:Id>",
                                "<a:LogicalName>", principalEntityName, "</a:LogicalName>",
                                "<a:Name i:nil='true' />",
                                "</b:Principal>",
                            "</c:value>",
                            "</a:KeyValuePairOfstringanyType>",
                        "</a:Parameters>",
                        "<a:RequestId i:nil='true' />",
                        "<a:RequestName>ModifyAccess</a:RequestName>",
                    "</request>"].join("");
        var async = !!callback;

        return doRequest(request, "Execute", async, function (resultXml) {
            var responseText = selectSingleNodeText(resultXml, "//ser:ExecuteResult");
            var result = crmXmlDecode(responseText);
            if (!async)
                return result;
            else
                callback(result);
            // ReSharper disable NotAllPathsReturnValue
        });
        // ReSharper restore NotAllPathsReturnValue
    };

    var revokeAccess = function (accessOptions, callback) {
        ///<summary>
        /// Sends synchronous/asynchronous request to do a revokeAccess request.
        ///</summary>
        ///<param name="accessOptions" type="Object">
        /// A JavaScript Object with properties corresponding to the revokeAccess Criteria
        /// that are valid for revokeAccess operations.
        /// accessOptions.targetEntityName is a string represents the name of the target entity
        /// accessOptions.targetEntityId is a string represents the GUID of the target entity
        /// accessOptions.revokeeEntityName is a string represents the name of the revokee entity
        /// accessOptions.revokeeEntityId is a string represents the GUID of the revokee entity
        /// </param>
        ///<param name="callback" type="Function">
        /// A Function used for asynchronous request. If not defined, it sends a synchronous request.
        /// </param>

        var targetEntityName = accessOptions.targetEntityName;
        var targetEntityId = accessOptions.targetEntityId;
        var revokeeEntityName = accessOptions.revokeeEntityName;
        var revokeeEntityId = accessOptions.revokeeEntityId;

        var request = ["<request i:type='b:RevokeAccessRequest' xmlns:a='http://schemas.microsoft.com/xrm/2011/Contracts' xmlns:b='http://schemas.microsoft.com/crm/2011/Contracts'>",
                        "<a:Parameters xmlns:c='http://schemas.datacontract.org/2004/07/System.Collections.Generic'>",
                            "<a:KeyValuePairOfstringanyType>",
                            "<c:key>Target</c:key>",
                            "<c:value i:type='a:EntityReference'>",
                                "<a:Id>", encodeValue(targetEntityId), "</a:Id>",
                                "<a:LogicalName>", targetEntityName, "</a:LogicalName>",
                                "<a:Name i:nil='true' />",
                            "</c:value>",
                            "</a:KeyValuePairOfstringanyType>",
                            "<a:KeyValuePairOfstringanyType>",
                            "<c:key>Revokee</c:key>",
                            "<c:value i:type='a:EntityReference'>",
                                "<a:Id>", encodeValue(revokeeEntityId), "</a:Id>",
                                "<a:LogicalName>", revokeeEntityName, "</a:LogicalName>",
                                "<a:Name i:nil='true' />",
                            "</c:value>",
                            "</a:KeyValuePairOfstringanyType>",
                        "</a:Parameters>",
                        "<a:RequestId i:nil='true' />",
                        "<a:RequestName>RevokeAccess</a:RequestName>",
                    "</request>"].join("");
        var async = !!callback;

        return doRequest(request, "Execute", async, function (resultXml) {
            var responseText = selectSingleNodeText(resultXml, "//ser:ExecuteResult");
            var result = crmXmlDecode(responseText);
            if (!async)
                return result;
            else
                callback(result);
            // ReSharper disable NotAllPathsReturnValue
        });
        // ReSharper restore NotAllPathsReturnValue
    };

    var retrievePrincipalAccess = function (accessOptions, callback) {
        ///<summary>
        /// Sends synchronous/asynchronous request to do a retrievePrincipalAccess request.
        ///</summary>
        ///<param name="accessOptions" type="Object">
        /// A JavaScript Object with properties corresponding to the retrievePrincipalAccess Criteria
        /// that are valid for retrievePrincipalAccess operations.
        /// accessOptions.targetEntityName is a string represents the name of the target entity
        /// accessOptions.targetEntityId is a string represents the GUID of the target entity
        /// accessOptions.principalEntityName is a string represents the name of the principal entity
        /// accessOptions.principalEntityId is a string represents the GUID of the principal entity
        /// </param>
        ///<param name="callback" type="Function">
        /// A Function used for asynchronous request. If not defined, it sends a synchronous request.
        /// </param>

        var targetEntityName = accessOptions.targetEntityName;
        var targetEntityId = accessOptions.targetEntityId;
        var principalEntityName = accessOptions.principalEntityName;
        var principalEntityId = accessOptions.principalEntityId;

        var request = ["<request i:type='b:RetrievePrincipalAccessRequest' xmlns:a='http://schemas.microsoft.com/xrm/2011/Contracts' xmlns:b='http://schemas.microsoft.com/crm/2011/Contracts'>",
                        "<a:Parameters xmlns:c='http://schemas.datacontract.org/2004/07/System.Collections.Generic'>",
                            "<a:KeyValuePairOfstringanyType>",
                            "<c:key>Target</c:key>",
                            "<c:value i:type='a:EntityReference'>",
                                "<a:Id>", encodeValue(targetEntityId), "</a:Id>",
                                "<a:LogicalName>", targetEntityName, "</a:LogicalName>",
                                "<a:Name i:nil='true' />",
                            "</c:value>",
                            "</a:KeyValuePairOfstringanyType>",
                            "<a:KeyValuePairOfstringanyType>",
                            "<c:key>Principal</c:key>",
                            "<c:value i:type='a:EntityReference'>",
                                "<a:Id>", encodeValue(principalEntityId), "</a:Id>",
                                "<a:LogicalName>", principalEntityName, "</a:LogicalName>",
                                "<a:Name i:nil='true' />",
                            "</c:value>",
                            "</a:KeyValuePairOfstringanyType>",
                        "</a:Parameters>",
                        "<a:RequestId i:nil='true' />",
                        "<a:RequestName>RetrievePrincipalAccess</a:RequestName>",
                    "</request>"].join("");
        var async = !!callback;

        return doRequest(request, "Execute", async, function (resultXml) {
            var result = selectSingleNodeText(resultXml, "//b:value");
            if (!async)
                return result;
            else
                callback(result);
            // ReSharper disable NotAllPathsReturnValue
        });
        // ReSharper restore NotAllPathsReturnValue
    };

    // Added in 1.4.1 for metadata retrieval
    // Inspired From Microsoft SDK code to retrieve Metadata using JavaScript
    // Copyright (C) Microsoft Corporation.  All rights reserved.
    var arrayElements = ["Attributes",
                            "ManyToManyRelationships",
                            "ManyToOneRelationships",
                            "OneToManyRelationships",
                            "Privileges",
                            "LocalizedLabels",
                            "Options",
                            "Targets"];

    var isMetadataArray = function (elementName) {
        for (var i = 0, ilength = arrayElements.length; i < ilength; i++) {
            if (elementName === arrayElements[i]) {
                return true;
            }
        }
        return false;
    };

    var getNodeName = function (node) {
        if (typeof (node.baseName) != "undefined") {
            return node.baseName;
        }
        else {
            return node.localName;
        }
    };

    var objectifyNode = function (node) {
        //Check for null
        if (node.attributes != null && node.attributes.length === 1) {
            if (node.attributes.getNamedItem("i:nil") != null && node.attributes.getNamedItem("i:nil").nodeValue === "true") {
                return null;
            }
        }

        //Check if it is a value
        if ((node.firstChild != null) && (node.firstChild.nodeType === 3)) {
            var nodeName = getNodeName(node);

            switch (nodeName) {
                //Integer Values
                case "ActivityTypeMask":
                case "ObjectTypeCode":
                case "ColumnNumber":
                case "DefaultFormValue":
                case "MaxValue":
                case "MinValue":
                case "MaxLength":
                case "Order":
                case "Precision":
                case "PrecisionSource":
                case "LanguageCode":
                    return parseInt(node.firstChild.nodeValue, 10);
                    // Boolean values
                case "AutoRouteToOwnerQueue":
                case "CanBeChanged":
                case "CanTriggerWorkflow":
                case "IsActivity":
                case "IsActivityParty":
                case "IsAvailableOffline":
                case "IsChildEntity":
                case "IsCustomEntity":
                case "IsCustomOptionSet":
                case "IsDocumentManagementEnabled":
                case "IsEnabledForCharts":
                case "IsGlobal":
                case "IsImportable":
                case "IsIntersect":
                case "IsManaged":
                case "IsReadingPaneEnabled":
                case "IsValidForAdvancedFind":
                case "CanBeSecuredForCreate":
                case "CanBeSecuredForRead":
                case "CanBeSecuredForUpdate":
                case "IsCustomAttribute":
                case "IsPrimaryId":
                case "IsPrimaryName":
                case "IsSecured":
                case "IsValidForCreate":
                case "IsValidForRead":
                case "IsValidForUpdate":
                case "IsCustomRelationship":
                case "CanBeBasic":
                case "CanBeDeep":
                case "CanBeGlobal":
                case "CanBeLocal":
                    return (node.firstChild.nodeValue === "true") ? true : false;
                    //OptionMetadata.Value and BooleanManagedProperty.Value and AttributeRequiredLevelManagedProperty.Value
                case "Value":
                    //BooleanManagedProperty.Value
                    if ((node.firstChild.nodeValue === "true") || (node.firstChild.nodeValue === "false")) {
                        return (node.firstChild.nodeValue === "true") ? true : false;
                    }
                    //AttributeRequiredLevelManagedProperty.Value
                    if (
                            (node.firstChild.nodeValue === "ApplicationRequired") ||
                            (node.firstChild.nodeValue === "None") ||
                            (node.firstChild.nodeValue === "Recommended") ||
                            (node.firstChild.nodeValue === "SystemRequired")
                        ) {
                        return node.firstChild.nodeValue;
                    }
                    else {
                        //OptionMetadata.Value
                        return parseInt(node.firstChild.nodeValue, 10);
                    }
                    // ReSharper disable JsUnreachableCode
                    break;
                    // ReSharper restore JsUnreachableCode
                    //String values
                default:
                    return node.firstChild.nodeValue;
            }

        }

        //Check if it is a known array
        if (isMetadataArray(getNodeName(node))) {
            var arrayValue = [];
            for (var iii = 0, tempLength = node.childNodes.length; iii < tempLength; iii++) {
                var objectTypeName;
                if ((node.childNodes[iii].attributes != null) && (node.childNodes[iii].attributes.getNamedItem("i:type") != null)) {
                    objectTypeName = node.childNodes[iii].attributes.getNamedItem("i:type").nodeValue.split(":")[1];
                }
                else {

                    objectTypeName = getNodeName(node.childNodes[iii]);
                }

                var b = objectifyNode(node.childNodes[iii]);
                b._type = objectTypeName;
                arrayValue.push(b);

            }

            return arrayValue;
        }

        //Null entity description labels are returned as <label/> - not using i:nil = true;
        if (node.childNodes.length === 0) {
            return null;
        }

        //Otherwise return an object
        var c = {};
        if (node.attributes.getNamedItem("i:type") != null) {
            c._type = node.attributes.getNamedItem("i:type").nodeValue.split(":")[1];
        }
        for (var i = 0, ilength = node.childNodes.length; i < ilength; i++) {
            if (node.childNodes[i].nodeType === 3) {
                c[getNodeName(node.childNodes[i])] = node.childNodes[i].nodeValue;
            }
            else {
                c[getNodeName(node.childNodes[i])] = objectifyNode(node.childNodes[i]);
            }

        }
        return c;
    };

    var retrieveAllEntitiesMetadata = function (entityFilters, retrieveIfPublished, callback) {
        ///<summary>
        /// Sends an synchronous/asynchronous RetrieveAllEntitieMetadata Request to retrieve all entities metadata in the system
        ///</summary>
        ///<returns>Entity Metadata Collection</returns>
        ///<param name="entityFilters" type="Array">
        /// The filter array available to filter which data is retrieved. Case Sensitive filters [Entity,Attributes,Privileges,Relationships]
        /// Include only those elements of the entity you want to retrieve in the array. Retrieving all parts of all entities may take significant time.
        ///</param>
        ///<param name="retrieveIfPublished" type="Boolean">
        /// Sets whether to retrieve the metadata that has not been published.
        ///</param>
        ///<param name="callBack" type="Function">
        /// The function that will be passed through and be called by a successful response.
        /// This function also used as an indicator if the function is synchronous/asynchronous
        ///</param>

        entityFilters = isArray(entityFilters) ? entityFilters : [entityFilters];
        var entityFiltersString = "";
        for (var iii = 0, templength = entityFilters.length; iii < templength; iii++) {
            entityFiltersString += encodeValue(entityFilters[iii]) + " ";
        }

        var request = [
                "<request i:type=\"a:RetrieveAllEntitiesRequest\" xmlns:a=\"http://schemas.microsoft.com/xrm/2011/Contracts\">",
                "<a:Parameters xmlns:b=\"http://schemas.datacontract.org/2004/07/System.Collections.Generic\">",
                "<a:KeyValuePairOfstringanyType>",
                "<b:key>EntityFilters</b:key>",
                "<b:value i:type=\"c:EntityFilters\" xmlns:c=\"http://schemas.microsoft.com/xrm/2011/Metadata\">" + encodeValue(entityFiltersString) + "</b:value>",
                "</a:KeyValuePairOfstringanyType>",
                "<a:KeyValuePairOfstringanyType>",
                "<b:key>RetrieveAsIfPublished</b:key>",
                "<b:value i:type=\"c:boolean\" xmlns:c=\"http://www.w3.org/2001/XMLSchema\">" + encodeValue(retrieveIfPublished.toString()) + "</b:value>",
                "</a:KeyValuePairOfstringanyType>",
                "</a:Parameters>",
                "<a:RequestId i:nil=\"true\" />",
                "<a:RequestName>RetrieveAllEntities</a:RequestName>",
            "</request>"].join("");

        var async = !!callback;
        return doRequest(request, "Execute", async, function (resultXml) {
            var response = selectNodes(resultXml, "//c:EntityMetadata");

            var results = [];
            for (var i = 0, ilength = response.length; i < ilength; i++) {
                var a = objectifyNode(response[i]);
                a._type = "EntityMetadata";
                results.push(a);
            }

            if (!async)
                return results;
            else
                callback(results);
            // ReSharper disable NotAllPathsReturnValue
        });
        // ReSharper restore NotAllPathsReturnValue
    };

    var retrieveEntityMetadata = function (entityFilters, logicalName, retrieveIfPublished, callback) {
        ///<summary>
        /// Sends an synchronous/asynchronous RetreiveEntityMetadata Request to retrieve a particular entity metadata in the system
        ///</summary>
        ///<returns>Entity Metadata</returns>
        ///<param name="entityFilters" type="String">
        /// The filter string available to filter which data is retrieved. Case Sensitive filters [Entity,Attributes,Privileges,Relationships]
        /// Include only those elements of the entity you want to retrieve in the array. Retrieving all parts of all entities may take significant time.
        ///</param>
        ///<param name="logicalName" type="String">
        /// The string of the entity logical name
        ///</param>
        ///<param name="retrieveIfPublished" type="Boolean">
        /// Sets whether to retrieve the metadata that has not been published.
        ///</param>
        ///<param name="callBack" type="Function">
        /// The function that will be passed through and be called by a successful response.
        /// This function also used as an indicator if the function is synchronous/asynchronous
        ///</param>

        entityFilters = isArray(entityFilters) ? entityFilters : [entityFilters];
        var entityFiltersString = "";
        for (var iii = 0, templength = entityFilters.length; iii < templength; iii++) {
            entityFiltersString += encodeValue(entityFilters[iii]) + " ";
        }

        var request = [
            "<request i:type=\"a:RetrieveEntityRequest\" xmlns:a=\"http://schemas.microsoft.com/xrm/2011/Contracts\">",
                "<a:Parameters xmlns:b=\"http://schemas.datacontract.org/2004/07/System.Collections.Generic\">",
                    "<a:KeyValuePairOfstringanyType>",
                        "<b:key>EntityFilters</b:key>",
                        "<b:value i:type=\"c:EntityFilters\" xmlns:c=\"http://schemas.microsoft.com/xrm/2011/Metadata\">", encodeValue(entityFiltersString), "</b:value>",
                    "</a:KeyValuePairOfstringanyType>",
                    "<a:KeyValuePairOfstringanyType>",
                        "<b:key>MetadataId</b:key>",
                        "<b:value i:type=\"c:guid\"  xmlns:c=\"http://schemas.microsoft.com/2003/10/Serialization/\">", encodeValue("00000000-0000-0000-0000-000000000000"), "</b:value>",
                    "</a:KeyValuePairOfstringanyType>",
                    "<a:KeyValuePairOfstringanyType>",
                        "<b:key>RetrieveAsIfPublished</b:key>",
                        "<b:value i:type=\"c:boolean\" xmlns:c=\"http://www.w3.org/2001/XMLSchema\">", encodeValue(retrieveIfPublished.toString()), "</b:value>",
                    "</a:KeyValuePairOfstringanyType>",
                    "<a:KeyValuePairOfstringanyType>",
                        "<b:key>LogicalName</b:key>",
                        "<b:value i:type=\"c:string\" xmlns:c=\"http://www.w3.org/2001/XMLSchema\">", encodeValue(logicalName), "</b:value>",
                    "</a:KeyValuePairOfstringanyType>",
                "</a:Parameters>",
                "<a:RequestId i:nil=\"true\" />",
                "<a:RequestName>RetrieveEntity</a:RequestName>",
                "</request>"].join("");

        var async = !!callback;

        return doRequest(request, "Execute", async, function (resultXml) {
            var response = selectNodes(resultXml, "//b:value");

            var results = [];
            for (var i = 0, ilength = response.length; i < ilength; i++) {
                var a = objectifyNode(response[i]);
                a._type = "EntityMetadata";
                results.push(a);
            }

            if (!async)
                return results;
            else
                callback(results);
            // ReSharper disable NotAllPathsReturnValue
        });
        // ReSharper restore NotAllPathsReturnValue
    };

    var retrieveAttributeMetadata = function (entityLogicalName, attributeLogicalName, retrieveIfPublished, callback) {
        ///<summary>
        /// Sends an synchronous/asynchronous RetrieveAttributeMetadata Request to retrieve a particular entity's attribute metadata in the system
        ///</summary>
        ///<returns>Entity Metadata</returns>
        ///<param name="entityLogicalName" type="String">
        /// The string of the entity logical name
        ///</param>
        ///<param name="attributeLogicalName" type="String">
        /// The string of the entity's attribute logical name
        ///</param>
        ///<param name="retrieveIfPublished" type="Boolean">
        /// Sets whether to retrieve the metadata that has not been published.
        ///</param>
        ///<param name="callBack" type="Function">
        /// The function that will be passed through and be called by a successful response.
        /// This function also used as an indicator if the function is synchronous/asynchronous
        ///</param>

        var request = [
                "<request i:type=\"a:RetrieveAttributeRequest\" xmlns:a=\"http://schemas.microsoft.com/xrm/2011/Contracts\">",
                "<a:Parameters xmlns:b=\"http://schemas.datacontract.org/2004/07/System.Collections.Generic\">",
                "<a:KeyValuePairOfstringanyType>",
                "<b:key>EntityLogicalName</b:key>",
                "<b:value i:type=\"c:string\" xmlns:c=\"http://www.w3.org/2001/XMLSchema\">", encodeValue(entityLogicalName), "</b:value>",
                "</a:KeyValuePairOfstringanyType>",
                "<a:KeyValuePairOfstringanyType>",
                "<b:key>MetadataId</b:key>",
                "<b:value i:type=\"ser:guid\"  xmlns:ser=\"http://schemas.microsoft.com/2003/10/Serialization/\">", encodeValue("00000000-0000-0000-0000-000000000000"), "</b:value>",
                "</a:KeyValuePairOfstringanyType>",
                "<a:KeyValuePairOfstringanyType>",
                "<b:key>RetrieveAsIfPublished</b:key>",
                "<b:value i:type=\"c:boolean\" xmlns:c=\"http://www.w3.org/2001/XMLSchema\">", encodeValue(retrieveIfPublished.toString()), "</b:value>",
                "</a:KeyValuePairOfstringanyType>",
                "<a:KeyValuePairOfstringanyType>",
                "<b:key>LogicalName</b:key>",
                "<b:value i:type=\"c:string\"   xmlns:c=\"http://www.w3.org/2001/XMLSchema\">", encodeValue(attributeLogicalName), "</b:value>",
                "</a:KeyValuePairOfstringanyType>",
                "</a:Parameters>",
                "<a:RequestId i:nil=\"true\" />",
                "<a:RequestName>RetrieveAttribute</a:RequestName>",
                "</request>"].join("");

        var async = !!callback;

        return doRequest(request, "Execute", async, function (resultXml) {
            var response = selectNodes(resultXml, "//b:value");
            var results = [];
            for (var i = 0, ilength = response.length; i < ilength; i++) {
                var a = objectifyNode(response[i]);
                results.push(a);
            }

            if (!async)
                return results;
            else
                callback(results);
            // ReSharper disable NotAllPathsReturnValue
        });
        // ReSharper restore NotAllPathsReturnValue
    };
}