<script src="https://code.jquery.com/jquery-2.1.4.min.js"></script>
<script type="text/javascript">
function AddToDelve() {

    var appWebUrl = window.location.protocol + "//" + window.location.host
                + _spPageContextInfo.webServerRelativeUrl;
    var board = $("#board").val();
    var data1 = {
        "signals": 
[{"Type":"Tag","TagName":"TEST",
  "DocumentUrl":"https://fancycoder-my.sharepoint.com/personal/natalliamakarevich_fancycoder_onmicrosoft_com/Documents/Document1.docx"},
 {"Type":"Follow","TagName":"TAG://PUBLIC/?NAME=TEST"}]
    };


    $.ajax({
        url: appWebUrl + '/_api/contextinfo',
        method: "POST",
        headers: { "Accept": "application/json; odata=verbose" }
    }).done(function (data) {
        var validRequestDigest = data.d.GetContextWebInformation.FormDigestValue;
       // debugger;
        

        var requestHeaders = {
            "Accept": "application/json;odata=verbose",
            "Access-Control-Allow-Origin": "https://fancycoder.sharepoint.com",
            "X-RequestDigest": validRequestDigest,
	    "X-Delve-ClientPlatform":"DelveWeb"
        };

        jQuery.ajax({
            url: appWebUrl + "_vti_bin/DelveApi.ashx/signals/batch",
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