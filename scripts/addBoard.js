<script src="https://code.jquery.com/jquery-2.1.4.min.js"></script>
<script type="text/javascript">
function AddToDelve(board, docLocation) {
    var host = window.location.protocol + "//" + window.location.host;
    var appWebUrl = host + _spPageContextInfo.webServerRelativeUrl;
    
    var data1 = {
        "signals": [{"Type":"Tag","TagName":board, "DocumentUrl" : docLocation},
        {"Type":"Follow","TagName":"TAG://PUBLIC/?NAME=" + board}]
    };


    $.ajax({
        url: appWebUrl + '/_api/contextinfo',
        method: "POST",
        headers: { "Accept": "application/json; odata=verbose" }
    }).done(function (data) {
        var validRequestDigest = data.d.GetContextWebInformation.FormDigestValue;
       
        var requestHeaders = {
            "Accept": "application/json;odata=verbose",            
            "X-RequestDigest": validRequestDigest,
            "X-Delve-ClientPlatform":"DelveWeb",
            "Access-Control-Allow-Origin": host
        };

        jQuery.ajax({
            url: appWebUrl + "/_vti_bin/DelveApi.ashx/signals/batch",
            type: "POST",
            data: JSON.stringify(data1.signals),
            contentType: "application/json;odata=verbose",
            headers: requestHeaders,
            success: function (data1) {
                console.log(data1);
            },
            error: function (jqxr, errorCode, errorThrown) {
                console.log(jqxr.responseText);
            }
        });
    });

}

</script>