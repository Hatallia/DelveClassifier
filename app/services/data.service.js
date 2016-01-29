(function () {
    'use strict';

    angular.module('officeAddin').service('dataService', ['sharePointUrl', '$http', '$q', dataService]);

    /**
     * Custom Angular service.
     */
    function dataService(sharePointUrl, $http, $q) {

        // public signature of the service
        return {
            getAllBoards: getBoards,
            getBoardDocuments: getBoardDocuments,
            getDocumentByUrl: getDocumentByUrl
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

        function getDocumentByUrl(docUrl) {
                var deferred = $q.defer();
            var searchQuery = "?QueryText=%27Path:" + docUrl + "%27&SelectProperties=%27Title,Path%27";
            $http({
                url: sharePointUrl + '/_api/search/query' + searchQuery,
                method: 'GET',
                headers: {
                    'Accept': 'application/json;odata=nometadata'
                }
            }).success(function (data) {
                var document = {
                    title: undefined,
                    url:docUrl
                };
                if (data.PrimaryQueryResult !== null) {
                    data.PrimaryQueryResult.RelevantResults.Table.Rows.forEach(function (row) {
                        var cells = row.Cells;
                        var url = getValueFromResults('Path', cells);
                        if (url === docUrl) {
                            document.title =  getValueFromResults('Title', cells);
                        }
                    });
                }
                deferred.resolve(document);

            }).error(function (err) {
                deferred.reject(err);
            });
            return deferred.promise;
        }

        function getBoardDocuments(board) {
            var deferred = $q.defer();
            var searchQuery = "?QueryText=%27*%27&Properties=%27GraphQuery:actor(" + board.docId + "\\,action\\:1045)%27&RankingModelId=%270c77ded8-c3ef-466d-929d-905670ea1d72%27";
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

                        var url = getValueFromResults('Path', cells);
                        /*if (url === null) {
                            url = getValueFromResults('ServerRedirectedURL', cells);
                        }*/

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

        function getBoards() {
            var deferred = $q.defer();
            var searchQuery = "?QueryText=%27Path:TAG://PUBLIC/?NAME=*%27&Properties=%27IncludeExternalContent:true%27&SelectProperties=%27DocId,Title,Path%27&RankingModelId=%270c77ded8-c3ef-466d-929d-905670ea1d72%27";
            $http({
                url: sharePointUrl + '/_api/search/query' + searchQuery,
                method: 'GET',
                headers: {
                    'Accept': 'application/json;odata=nometadata'
                }
            }).success(function (data) {
                var boards = [];
                if (data.PrimaryQueryResult !== null) {
                    data.PrimaryQueryResult.RelevantResults.Table.Rows.forEach(function (row) {
                        var cells = row.Cells;
                        var url = sharePointUrl + '/_layouts/15/me.aspx?b=' + getValueFromResults('Title', cells);
                        boards.push({
                            url: url,
                            title: getValueFromResults('Title', cells),
                            docId: getValueFromResults('DocId', cells)
                        });
                    });

                    boards.sort(function (a, b) {
                        return a.title.toLowerCase().localeCompare(b.title.toLowerCase());
                    });
                }
                //Now we  will try to get where the document is added to.
                deferred.resolve(boards);
            }).error(function (err) {
                deferred.reject(err);
            });
            return deferred.promise;
        }
    }    
})();