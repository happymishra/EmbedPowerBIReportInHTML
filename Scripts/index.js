(function () {
    var resourceUri = "https://analysis.windows.net/powerbi/api",
     iframe = document.getElementById('iframeId'),
     reportURL = "https://msit.powerbi.com/reportEmbed?reportId=",
     authContext,
     isCallback,
     user,
     regex,
     results,
     reportIdValue,
     errorMsg,
     regexS,
     regex,
     results,
     listUrlForReportId;

    window.config = {
        instance: 'https://login.microsoftonline.com/',
        tenant: 'microsoft.onmicrosoft.com', //Enter tenant Name e.g. microsoft.onmicrosoft.com
        clientId: '{{clientId}}', //Enter your app Client ID created in Azure Portal
        redirectUri: "{{redirectUri}}",
        cacheLocation: 'localStorage', // enable this for IE, as sessionStorage does not work for localhost.
    };

    authContext = new AuthenticationContext(config);
    isCallback = authContext.isCallback(window.location.hash);
    authContext.handleWindowCallback();

    if (isCallback && !authContext.getLoginError()) {
        window.location = authContext._getItem(authContext.CONSTANTS.STORAGE.LOGIN_REQUEST);
    }

    user = authContext.getCachedUser();
    if (!user) {
        authContext.login();
    }

    //get modelid passed in the Url
    reportIdValue = getModelIdFromUrl('reportId', location.search);
    if (null === reportIdValue) {
        errorMsg("Invalid Url");
        return;
    }

    // Power BI reports URL are stored in a SharePoint list
    // Fetch the report URL from the SharePoint list using the reportId passed in the url
    listUrlForReportId = "/_api/Web/Lists/GetByTitle('PowerBIReportURLSharePointList')/Items?$filter=Title eq '" + reportIdValue + "'";
    getReportIdAndCallEmbedReport(listUrlForReportId);

    function getReportIdFromUrl(name, url) {
        if (!url) url = location.href;
        name = name.replace(/[\[]/, "\\\[").replace(/[\]]/, "\\\]");
        regexS = "[\\?&]" + name + "=([^&#]*)";
        regex = new RegExp(regexS, "i");
        results = regex.exec(url);
        return results == null ? null : results[1];
    }

    function errorMsg(message) {
        iframe.style.display = "none";
        $('#invalidUrl').text("Error: " + message);
    }

    function getReportIdAndCallEmbedReport(url) {
        $.ajax({
            url: _spPageContextInfo.webAbsoluteUrl + url,
            type: "GET",
            headers: {
                "accept": "application/json;odata=verbose",
            },
            success: function (data) {
                if (0 < data.d.results.length) {
                    reportURL = reportURL + data.d.results[0].Report_x0020_Id;
                    embedReport(reportURL);
                }
                else {
                    errorMsg("Report for the model does not exists");
                }
            },
            error: function (error) {
                console.log(JSON.stringify(error));
            }
        });
    }

    function embedReport(reportURL) {
        iframe.src = reportURL;
        iframe.onload = postActionLoadTile;
        return false;
    }

    // Post the authentication token to the IFrame.
    function postActionLoadTile() {
        authContext.acquireToken(resourceUri, function (error, token) { //Add Resource ID here e.g. microsoft.onmicrosoft.com

            // Handle ADAL Error
            if (error || !token) {
                console.log('ADAL Error Occurred: ' + error);
                return;
            }
            console.log("token" + token);
            if ("" == token)
                return;

            // construct the post message structure
            var m = { action: "loadReport", accessToken: token };
            message = JSON.stringify(m);
            iframe.contentWindow.postMessage(message, "*");
        });
    }
}());