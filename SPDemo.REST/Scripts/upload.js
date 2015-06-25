'use strict';

var hostWebUrl;
var appWebUrl;
var context = SP.ClientContext.get_current();

function initializePage() {
    alert("initialize");
    $.ajaxSetup({ cache: false, crossDomain: true });
    $.support.cors = true;
    hostWebUrl = decodeURIComponent(getQueryStringParameter("SPHostUrl"));
    appWebUrl = decodeURIComponent(getQueryStringParameter("SPAppWebUrl"));
    setupControls();
}
function setupControls() {
    alert("setup");
    $("#buttonSubmit").click(function () {
        uploadFile();
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

//Document Upload
function uploadFile() {
    alert("upload");
    var fileInput = $('#inputFile')[0];

    if (fileInput == null || fileInput.length == 0) {
        alert("Please select a file.");
        return;
    } else {
        try {
            var pathArray = $("#inputFile").val().split("\\");
            var fileName = pathArray[pathArray.length - 1];
            var reader = new FileReader();

            reader.onload = function (result) {
                var fileData = '';
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
                    "content-length": reader.byteLength,
                    "content-type": "application/json;odata=verbose",
                    "X-FORMS_BASED_AUTH_ACCEPTED": "f"
                },
                body: reader,
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