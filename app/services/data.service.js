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
      getDocuments: getDocuments,
      getAllBoards: getAllBoards
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
 
    searchQuery = "?QueryText=%27Path:TAG://PUBLIC/?NAME=*%27&Properties=%27IncludeExternalContent:true%27&SelectProperties=%27DocId,Title,Path%27&RankingModelId=%270c77ded8-c3ef-466d-929d-905670ea1d72%27";

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
    
  





 function getAllBoards() {
      var deferred = $q.defer();

   var searchQuery = "?QueryText=%27Path:TAG://PUBLIC/?NAME=*%27&Properties=%27IncludeExternalContent:true%27&SelectProperties=%27DocId,Title,Path%27&RankingModelId=%270c77ded8-c3ef-466d-929d-905670ea1d72%27";

    $http({
        url: sharePointUrl + '/_api/search/query' + searchQuery,
        method: 'GET',
        headers: {
          'Accept': 'application/json;odata=nometadata'
        }
      }).success(function (data) {

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
              title: getValueFromResults('Title', cells)              
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