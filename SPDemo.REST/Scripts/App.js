'use strict';

var hostWebUrl;
var appWebUrl;
var context = SP.ClientContext.get_current();
var dispFormUrl;

function initializePage() {
    $.ajaxSetup({ cache: false, crossDomain: true });
    $.support.cors = true;
    hostWebUrl = decodeURIComponent(getQueryStringParameter("SPHostUrl"));
    appWebUrl = decodeURIComponent(getQueryStringParameter("SPAppWebUrl"));
    setupControls();
}
function setupControls() {
    $("#tabstrip").kendoTabStrip({
        animation: {
            open: {
                effects: "fadeIn"
            }
        }
    });
    $("#restOptions").kendoComboBox({
        dataTextField: "text",
        dataValueField: "value",
        select: function (e) {
            var dataItem = this.dataItem(e.item.index());
            $("#demoCommand_Text").html(dataItem.value);
        }
    });
    $("#buttonSubmit").click(function () {
        loadResults();
        return false;
    });
    $("#fileSubmit").click(function () {
        uploadFile();
        return false;
    });
    $("#siteSubmit").click(function () {
        createSite();
        return false;
    });
    $("#searchSubmit").click(function () {
        searchQuery();
        return false;
    });
    $("#socialSubmit").click(function () {
        getFollowed();
        return false;
    });
    $("#sharingSubmit").click(function () {
        shareDocument();
        return false;
    });
}
function getQueryStringParameter(paramToRetrieve) {
    var params =
                document.URLUnencoded.split("?")[1].split("&");
    var strParams = "";
    for (var i = 0; i < params.length; i = i + 1) {
        var singleParam = params[i].split("=");
        if (singleParam[0] == paramToRetrieve)
            return singleParam[1];
    }
}

//List Items
function loadResults() {
    var select = $("#restOptions").data("kendoComboBox");

    try {
        var results;
        var executor = new SP.RequestExecutor(appWebUrl);
        var query = select.value();
        var reqUrl;
        if (select.text() == 'Dynamic View') {
            var qOptions = select.value().split(';');
            var qControl = qOptions[0];
            var qWeb = qOptions[1];
            var qList = qOptions[2];
            var qView = qOptions[3];

            var listUrl = appWebUrl + "/_api/SP.AppContextSite(@target)/web/lists/getbytitle('" + qList + "')?$select=defaultdisplayformurl&@target='" + hostWebUrl + "'";

            executor.executeAsync({
                url: listUrl,
                method: "GET",
                headers: { "Accept": "application/json;odata=verbose" },
                success: function (data) {
                    var objData = JSON.parse(data.body);
                    results = objData.d.results;
                    if (!results) {
                        var r = objData.d;
                        var o = '[' + JSON.stringify(r) + ']';
                        results = $.parseJSON(o);
                        var a = $('<a>', { href: hostWebUrl })[0];
                        dispFormUrl = a.protocol + "//" + a.hostname + results[0].DefaultDisplayFormUrl;
                        getListView(qControl, qWeb, qList, qView, '1000');
                    }
                },
                error: function (result, code, message) {
                    alert(message);
                }
            });
        } else {

            if (query.indexOf('?') != -1) {
                reqUrl = appWebUrl + "/_api/SP.AppContextSite(@target)" + select.value() + "&@target='" + hostWebUrl + "'";
            } else {
                reqUrl = appWebUrl + "/_api/SP.AppContextSite(@target)" + select.value() + "?@target='" + hostWebUrl + "'";
            }

            executor.executeAsync({
                url: reqUrl,
                method: "GET",
                headers: { "Accept": "application/json;odata=verbose" },
                success: function (data) {
                    var objData = JSON.parse(data.body);
                    results = objData.d.results;
                    if (!results) {
                        var r = objData.d;
                        var o = '[' + JSON.stringify(r) + ']';
                        results = $.parseJSON(o);
                    }

                    $("#demoGrid").empty();
                    $("#demoGrid").kendoGrid({
                        theme: $(document).data("kendoSkin") || "Metro",
                        dataSource: {
                            data: results,
                            total: function (data) {
                                return data.length;
                            },
                            pageSize: 20
                        },
                        filterable: true,
                        sortable: true,
                        pageSize: 20,
                        pageable: true,
                        columns: [
                        { field: "ID", title: "ID", attributes: { style: "white-space: nowrap;text-align:left;" } },
                        { field: "Title", title: "Title", attributes: { style: "white-space: nowrap;text-align:left;" } },
                        { field: "State", title: "State", attributes: { style: "white-space: nowrap;" } },
                        { field: "Code", title: "Code", attributes: { style: "white-space: nowrap" } },
                        { field: "Industry", title: "Industry", attributes: { style: "white-space: nowrap" } },
                        { field: "Year", title: "Year", ttributes: { style: "white-space: nowrap" } },
                        { field: "Amount", title: "Amount", attributes: { style: "white-space: nowrap" }, format: "{0:c}" }]
                    });
                },
                error: function (result, code, message) {
                    alert(message);
                }
            });
        }
    } catch (err) {
        alert(err.message);
    }   
}

function getListView(control, web, list, view, limit) {
    var results;
    var executor = new SP.RequestExecutor(appWebUrl);
    var viewFields = [];

    var viewUrl = appWebUrl + "/_api/SP.AppContextSite(@target)/web/lists/getbytitle('" + list + "')/views('" + view + "')/viewfields?@target='" + web + "'";

    executor.executeAsync({
        url: viewUrl,
        method: "GET",
        headers: { "Accept": "application/json;odata=verbose" },
        success: function (data) {
            var objData = JSON.parse(data.body);
            results = objData.d.results;
            if (!results) {
                var r = objData.d;
                var o = '[' + JSON.stringify(r) + ']';
                results = $.parseJSON(o);
                $.each(results[0].Items.results, function (i, val) {
                    viewFields.push(val + "/FieldValuesAsHtml");
                });
                if ($.inArray("ID/FieldValuesAsHtml", viewFields) == -1) {
                    viewFields.push("ID/FieldValuesAsHtml");
                }
                createViewTable(control, web, list, view, viewFields, limit);
            }
        },
        error: function (result, code, message) {
            alert(message);
        }
    });
}

function createViewTable(control, web, list, view, fields, limit) {
    try {
        var results;
        var executor = new SP.RequestExecutor(appWebUrl);
        var cols = null;

        var columnSchema = [];
        for (var i = 0; i < fields.length; i++) {
            var fieldVals = fields[i].split('/');
            var fieldVal = fieldVals[0];
            columnSchema.push({ field: fieldVal, title: fieldVal, headerAttributes: { style: "font-weight:bold;" }, attributes: { style: "text-align:left;font-weight:normal;" } });
        }

        var reqUrl = appWebUrl + "/_api/SP.AppContextSite(@target)/web/lists/getbytitle('" + list + "')/items?$select=" + fields.join() + "&$top=" + limit + "&@target='" + web + "'";

        $("#demoCommand_Text").html(reqUrl);

        executor.executeAsync({
            url: reqUrl,
            method: "GET",
            headers: { "Accept": "application/json;odata=verbose" },
            success: function (data) {
                var objData = JSON.parse(data.body);
                results = objData.d.results;
                if (!results) {
                    var r = objData.d;
                    var o = '[' + JSON.stringify(r) + ']';
                    results = $.parseJSON(o);
                }                

                $("#" + control).empty();
                $("#" + control).kendoGrid({
                    dataSource: {
                        data: results,
                        total: function (data) {
                            return data.length;
                        },
                        schema: {
                            parse: function (response) {
                                $.each(response, function (idx, elem) {
                                    $.each(elem, function (key, val) {
                                        try {
                                            if (key != "__metadata") {
                                                var valString = val.toString();
                                                if (valString.indexOf(':') != -1 && valString.indexOf('-') != -1 && valString.indexOf('T') != -1 && valString.indexOf('Z') != -1) {
                                                    var dt = new Date(val);
                                                    var d = dt.format('mm/dd/yyyy hh:mm:ss TT');
                                                    elem[key] = d;
                                                }                                            }
                                        } catch (err) {}
                                    });
                                });
                                return response
                            }
                        },
                        pageSize: 25,
                        batch: true
                    },
                    filterable: true,
                    sortable: true,
                    scrollable: false,
                    selectable: "row",
                    pageSize: 25,
                    pageable: { pageSizes: [25, 50, 100] },
                    columns: columnSchema,
                    change: function (arg) {
                        $.map(this.select(), function (item) {
                            var grid = $("#" + control).data("kendoGrid");
                            var data = grid.dataItem($(item));
                            var linkUrl = dispFormUrl + "?ID=" + data.ID;
                            window.location.href = linkUrl;
                        });
                    },
                });
            },
            error: function (result, code, message) {
                alert(message);
            }
        });
    } catch (err) {
        alert(err.message);
    }
}

// File Upload
function uploadFile() {
    var fileInput = $('#inputFile')[0];

    if (fileInput == null || fileInput.length == 0) {
        alert("Please select a file.");
        return;
    } else {
        try {
            var pathArray = $("#inputFile").val().split("\\");
            var fileName = pathArray[pathArray.length - 1];
            var reader = new FileReader();
            var fileData = '';

            reader.onload = function (result) {                
                var byteArray = new Uint8Array(result.target.result)
                for (var i = 0; i < byteArray.byteLength; i++) {
                    fileData += String.fromCharCode(byteArray[i])
                }
            };
            reader.readAsArrayBuffer(fileInput.files[0]);

            var reqUrl = appWebUrl + "/_api/SP.AppContextSite(@TargetSite)/web/lists/getByTitle('Documents')/RootFolder/Files/add(url=@TargetFileName,overwrite='true')?" +
                "@TargetSite='" + hostWebUrl + "'" +
                "&@TargetFileName='" + fileName + "'";

            var executor = new SP.RequestExecutor(appWebUrl);

            executor.executeAsync({
                url: reqUrl,
                method: "POST", 
                headers: {
                    "accept": "application/json;odata=verbose",
                    "X-RequestDigest": $("#__REQUESTDIGEST").val(),
                },
                contentType: "application/json;odata=verbose",
                body: fileData,
                binaryStringRequestBody: true,
                success: function (data) {
                    alert("File uploaded successfully.");
                },
                error: function (result, code, message) {
                    alert(message);
                }
            });

        } catch (err) {
            alert(err.message);
        }
    }
}

//Site Provisioning
function createSite() {
    try {
        $("#siteStatus").html("Please wait...")
        var siteName = $('#inputSite').val();
        var reqUrl = appWebUrl + "/_api/SP.AppContextSite(@target)/web/webinfos/add?@target='" + hostWebUrl + "'"; "";
        var executor = new SP.RequestExecutor(appWebUrl);
        var data = JSON.stringify({
            'parameters': {
                '__metadata': {
                    'type': 'SP.WebInfoCreationInformation'
                },
                'Url': siteName.toLowerCase().replace(" ",""),
                'Title': siteName,
                'Description': '',
                'Language': 1033,
                'WebTemplate': 'STS',
                'UseUniquePermissions': false
            }
        });

        executor.executeAsync({
            url: reqUrl,
            method: "POST",
            headers: {
                "accept": "application/json;odata=verbose",
                "content-type": "application/json;odata=verbose",
                "X-RequestDigest": $("#__REQUESTDIGEST").val(),
            },
            body: data,
            success: function (data) {
                $("#siteStatus").html("Provisioning succeeded for site '" + siteName + "'.").css("color", "green");
            },
            error: function (result, code, message) {
                $("#siteStatus").html("Provisioning failed.").css("color", "red");
            }
        });
    } catch (err) {
        alert(err.message);
    }

}

//Search
function searchQuery() {
    try {
        var keywords = $("#inputSearch").val();
        var reqUrl = appWebUrl + "/_api/search/query?queryText='" + keywords + "'";
        var executor = new SP.RequestExecutor(appWebUrl);
        executor.executeAsync({
            url: reqUrl,
            method: "GET",
            headers: { "Accept": "application/json;odata=verbose" },
            success: function (data) {
                var objData = JSON.parse(data.body);
                var results = objData.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results;
                if (results.length == 0) {
                    $('#searchOutput').html('<p>No results found.</p>');
                } else {
                    var resultsHtml = '';
                    $.each(results, function (index, result) {
                        resultsHtml += "<a target='_blank' href='" + result.Cells.results[6].Value + "'>" + result.Cells.results[3].Value + "</a><br/><span style='font-weight:normal;font-size:10pt;'>Author: " + result.Cells.results[4].Value + "<br/>" + result.Cells.results[10].Value + "</span><br/><br/>";
                    });
                    $('#searchOutput').html(resultsHtml);
                }
            },
            error: function (result, code, message) {
                alert(message);
            }
        });
    } catch (err) {
        alert(err.message);
    }
}


//User Profiles
function getFollowed() {
    try {
        var reqUrl = appWebUrl + "/_api/social.following/my/followed(types=14)?$orderby=ActorType";
        var executor = new SP.RequestExecutor(appWebUrl);
        executor.executeAsync({
            url: reqUrl,
            method: "GET",
            headers: { "Accept": "application/json;odata=verbose" },
            success: function (data) {
                var objData = JSON.parse(data.body);
                var results = objData.d.Followed.results;
                if (results.length == 0) {
                    $('#socialOutput').html('<p>No results found.</p>');
                } else {
                    var resultsHtml = '';
                    $.each(results, function (index, result) {
                        resultsHtml += "<a target='_blank' href='" + result.Uri + "'>" + result.Name + "</a><br/><br/>";
                    });
                    $('#socialOutput').html(resultsHtml);
                }
            },
            error: function (result, code, message) {
                alert(message);
            }
        });
    } catch (err) {
        alert(err.message);
    }
}

//Sharing
function shareDocument() {
    try {

        var ctxWeb = context.get_web();
        var currentUser = ctxWeb.get_currentUser();
        context.load(currentUser);
        context.executeQueryAsync(
            function () {
                var userId = currentUser.get_loginName();
                var docUrl = $('#inputFileUrl').val();
                var executor = new SP.RequestExecutor(appWebUrl);
                var restSource = appWebUrl + "/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerResolveUser";

                var restData = JSON.stringify({
                    'queryParams': {
                        '__metadata': {
                            'type': 'SP.UI.ApplicationPages.ClientPeoplePickerQueryParameters'
                        },
                        'AllowEmailAddresses': true,
                        'AllowMultipleEntities': false,
                        'AllUrlZones': false,
                        'MaximumEntitySuggestions': 50,
                        'PrincipalSource': 15,
                        'PrincipalType': 1,
                        'QueryString': userId.toString()
                        //'Required':false,
                        //'SharePointGroupID':null,
                        //'UrlZone':null,
                        //'UrlZoneSpecified':false,
                        //'Web':null,
                        //'WebApplicationID':null
                    }
                });
                executor.executeAsync({
                    url: restSource,
                    method: "POST",
                    headers: {
                        "accept": "application/json;odata=verbose",
                        "content-type": "application/json;odata=verbose",
                        "X-RequestDigest": $("#__REQUESTDIGEST").val(),
                    },
                    body: restData,
                    success: function (data) {
                        var body = JSON.parse(data.body);
                        var results = body.d.ClientPeoplePickerResolveUser;
                        if (results.length > 0) {
                            var reqUrl = appWebUrl + "/_api/SP.Web.ShareObject";
                            var executor = new SP.RequestExecutor(appWebUrl);
                            var data = JSON.stringify({
                                    "url": docUrl,
                                    "peoplePickerInput": '[' + results + ']',
                                    "roleValue":  "1073741827",
                                    "groupId":  0,
                                    "propagateAcl":  false,
                                    "sendEmail":  true,
                                    "includeAnonymousLinkInEmail":  true,
                                    "emailSubject":  "Sharing Test",
                                    "emailBody":  "This is a Sharing Test."
                            });

                            executor.executeAsync({
                                url: reqUrl,
                                method: "POST",
                                headers: {
                                    "accept": "application/json;odata=verbose",
                                    "content-type": "application/json;odata=verbose",
                                    "X-RequestDigest": $("#__REQUESTDIGEST").val(),
                                },
                                body: data,
                                success: function (data) {
                                    $("#sharingOutput").html("Sharing succeeded for '" + docUrl + "'.").css("color", "green");
                                },
                                error: function (result, code, message) {
                                    $("#sharingOutput").html(message).css("color", "red");
                                }
                            });
                        }
                        
                    },
                    error: function (result, code, message) {
                        $("#sharingOutput").html(message).css("color", "red");
                    }
                });

            },
            function (sender, args) {
                $("#sharingOutput").html(args.get_message()).css("color", "red");
            }
        );
    } catch (err) {
        alert(err.message);
    }

}

