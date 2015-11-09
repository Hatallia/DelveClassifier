(function () {
  'use strict';

  angular.module('officeAddin')
    .service('dataService', ['sharePointUrl', '$http', '$q', dataService]);

  /**
   * Custom Angular service.
   */
  function dataService(sharePointUrl, $http, $q) {
    
    // public signature of the service
    return {
      getDocuments: getDocuments
    };

    /** *********************************************************** */

    function getValueFromResults(key, results) {
      var value = '';

      if (results !== null &&
        results.length > 0 &&
        key !== null) {
        for (var i = 0; i < results.length; i++) {
          var resultItem = results[i];

          if (resultItem.Key === key) {
            value = resultItem.Value;
            break;
          }
        }
      }

      return value;
    }
    
    function getDocuments(query) {
      var deferred = $q.defer();
 var searchQuery = "?QueryText=%27Path:TAG://PUBLIC/?NAME=" + query + "%27&Properties=%27IncludeExternalContent:true%27&SelectProperties=%27DocId,Title,Path%27&RankingModelId=%270c77ded8-c3ef-466d-929d-905670ea1d72%27";
    //  var searchQuery = "?querytext='" + query + " isdocument:1'&SelectProperties='HitHighlightedSummary,LastModifiedTime,Path,SPWebUrl,ServerRedirectedURL,SiteTitle,Title'&RowLimit=5&StartRow=0";
 //var searchQuery = "?QueryText=%27Path:TAG://PUBLIC/?NAME=*%27&Properties=%27IncludeExternalContent:true%27&SelectProperties=%27DocId,Title,Path%27&RankingModelId=%270c77ded8-c3ef-466d-929d-905670ea1d72%27";

      

         
// var signal = [{"Type":"Tag","TagName":"TEST",
//   "DocumentUrl":"https://fancycoder-my.sharepoint.com/personal/natalliamakarevich_fancycoder_onmicrosoft_com/Documents/Document1.docx"},
//  {"Type":"Follow","TagName":"TAG://PUBLIC/?NAME=TEST"}];


//       $http({
//         url: "https://fancycoder-my.sharepoint.com/_vti_bin/DelveApi.ashx/signals/batch",
//         method: 'POST',
//         headers: {
//           'Accept': 'application/json;odata=nometadata',
//           'Content-Type':'application/json;odata=verbose',
//           'Host':'fancycoder-my.sharepoint.com',
//           'Origin':'https://fancycoder-my.sharepoint.com',

//           "Access-Control-Allow-Origin":"*",
//           'X-RequestDigest':query,
//           'X-Delve-ClientPlatform':'DelveWeb'

//         },
//         data:JSON.stringify(signal)
      $http({
        url: sharePointUrl + '/_api/search/query' + searchQuery,
        method: 'GET',
        headers: {
          'Accept': 'application/json;odata=nometadata'
        }
      }).success(function (data) {

      	///deferred.resolve([{url:"",summary:"",title:data, siteUrl:"",siteTitle:""}]);

        var documents = [];

        if (data.PrimaryQueryResult !== null) {
          data.PrimaryQueryResult.RelevantResults.Table.Rows.forEach(function (row) {
            var cells = row.Cells;

            var url = getValueFromResults('ServerRedirectedURL', cells);
            if (url === null) {
              url = getValueFromResults('Path', cells);
            }

            documents.push({
              url: url,
              title: getValueFromResults('Title', cells),
              summary: getValueFromResults('HitHighlightedSummary', cells).replace(/<(\/)?c\d>/g, '<$1mark>').replace(/<ddd\/>/g, ''),
              siteUrl: getValueFromResults('SPWebUrl', cells),
              siteTitle: getValueFromResults('SiteTitle', cells)
            });
          });
        }

        deferred.resolve(documents);
      }).error(function (err) {
        deferred.reject(err);
      });

      return deferred.promise;
    }
    
  }
})();