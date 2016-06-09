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
            getDocumentByUrl: getDocumentByUrl,
			getFormDigest: getFormDigest
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
            var searchQuery = "?QueryText=%27*%27&Properties=%27GraphQuery:actor(" + board.docId + "\\,action\\:1045)%27&RankingModelId=%270c77ded8-c3ef-466d-929d-905670ea1d72%27&SelectProperties=%27Title,Path,ServerRedirectedURL%27";
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
                        var onlineViewUrl = getValueFromResults('ServerRedirectedURL', cells);

                        documents.push({
                            url: url,
                            onlineUrl: onlineViewUrl,
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
		
		function getFormDigest(){	
			$http({
                //url: sharePointUrl + '/_api/search/query' + searchQuery,
				url: sharePointUrl + "/_api/contextinfo",
                method: 'POST',
                headers: {
                    "Accept": "application/json; odata=verbose", 
					"Content-Type":"application/json;odata=verbose"
                }
            }).success(function (data) {
				console.log(data);
				var digest = data.d.GetContextWebInformation.FormDigestValue;
                console.log(digest);
				addToBoardSignal(digest, 0);
            }).error(function (err) {
				console.log("Error");
                console.log(err);
            });
			
			
			/*$.ajax({url: "https://alexepam-my.sharepoint.com/_api/contextinfo", 
			   header: {
					"accept": "application/json; odata=verbose", 
					"content-type":"application/json;odata=verbose"}, 
					type: "POST", 
					contentType: "application/json;charset=utf-8"
			 }).done(function(d) {
					 console.log(d)
				  });*/
		}
		
		function getCookies(flag){	
			if (!flag) return;
			$http({
				url: sharePointUrl + "/_forms/default.aspx?wa=wsignin1.0",
                method: 'POST',
                headers: {
                    "Accept": "application/json; odata=verbose", 
					"Content-Type":"application/x-www-form-urlencoded"
                }
            }).success(function (data) {
				console.log(data);
            }).error(function (err) {
				console.log("Error");
                console.log(err);
            });
			
			
			/*$.ajax({url: "https://alexepam-my.sharepoint.com/_api/contextinfo", 
			   header: {
					"accept": "application/json; odata=verbose", 
					"content-type":"application/json;odata=verbose"}, 
					type: "POST", 
					contentType: "application/json;charset=utf-8"
			 }).done(function(d) {
					 console.log(d)
				  });*/
		}
		
		function addToBoard(digest){	
			//getCookies(true);
			var hds = {
				"Accept": "*/*", 
				//"Accept-Encoding": "gzip, deflate",
				"Content-Type":"application/json",
				"X-Delve-ClientPlatform": "DelveWeb",
				"X-RequestDigest": digest
			};			
			var dt = [{
				"Type":"Tag",
				"DocumentUrl":"https://alexepam-my.sharepoint.com/personal/aliaksandr_alexepam_onmicrosoft_com/Documents/Document8.docx",
				"TagName":"Board 002"
				},
				{
					"Type":"Follow",
					"TagName":"TAG://PUBLIC/?NAME=BOARD+002"
				}
			];
			//$http.defaults.withCredentials = true;
			$http({
				url: sharePointUrl + "/_vti_bin/DelveApi.ashx/signals/batch?flights=%27PulseWebFallbackCards,PulseWebStoryCards,PulseWebVideoCards,PulseWebContentTypeFilter%27",
                method: 'POST',
                headers: hds,
				data: dt
            }).success(function (data) {
				console.log(data);                
            }).error(function (err) {
				console.log("Error");
                console.log(err);
            });
		}
		
		function addToBoardSignal(digest, ch){	
			window.signalChoice = ch;
			var signalData = {        
			   "signals":[
				  {
					 "Actor":{
						"Id":"Aliaksandr@alexepam.onmicrosoft.com"//null
					 },
					 "Action":{
						"ActionType":"Tag",
						"UserTime": new Date().toISOString(),
						"Properties":[
							{
								"Key":"TagAction",
								"Value":"Add",
								"ValueType":"Edm.String"
							},
							{
								"Key":"TagName",
								"Value":"Board 002",
								"ValueType":"Edm.String"
							}
						]
					 },
					 "Item":{
						"Id":"https://alexepam-my.sharepoint.com/personal/aliaksandr_alexepam_onmicrosoft_com/Documents/Document8.docx"
					 },
					 "Source":"PulseWeb"
				  },
				  {
					 "Actor":{
						"Id":"Aliaksandr@alexepam.onmicrosoft.com"//null
					 },
					 "Action":{
						"ActionType":"Follow",
						"UserTime": new Date().toISOString(),
						"Properties":[
							{
								"Key":"ActionVerb",
								"Value":"Follow",
								"ValueType":"Edm.String"
							}
						]
					 },
					 "Item":{
						"Id":"TAG://PUBLIC/?NAME=BOARD+002"
					 },
					 "Source":"PulseWeb"
				  }
			   ]
			};
					
			var requestHeaders = {
				"Accept": "application/json;odata=verbose",
				"X-RequestDigest": digest
			};
			
			switch (window.signalChoice){
				case 0:{
					jQuery.ajax({
						url: sharePointUrl + "/_api/signalstore/signals",
						type: "POST",
						data: JSON.stringify(signalData),
						contentType: "application/json;odata=verbose",
						headers: requestHeaders,
						success: function (data) {
							console.log(data);
						},
						error: function (jqxr, errorCode, errorThrown) {
							console.log(jqxr.responseText);
						}
					});		
					break;
				}
				case 1:{
					jQuery.ajax({
						url: sharePointUrl + "/_api/signalstore/signals",
						type: "POST",
						data: JSON.stringify(signalData),
						contentType: "application/json;odata=verbose",
						headers: requestHeaders,
						crossDomain: true,
						//dataType: 'json',
						xhrFields: {
							withCredentials: true
						},	
						success: function (data) {
							console.log(data);
						},
						error: function (jqxr, errorCode, errorThrown) {
							console.log(jqxr.responseText);
						}
					});		
					break;
				}
				case 2: {
					$http({
						url: sharePointUrl + "/_api/signalstore/signals",
						method: 'POST',
						data: JSON.stringify(signalData),
						contentType: "application/json;odata=verbose",
						headers: requestHeaders
						}).success(function (data) {
							console.log(data);
						}).error(function (err) {
							console.log("Error");
							console.log(err);
					});
					break;
				}
			}		
		}
    }    
})();