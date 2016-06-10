(function () {
    'use strict';

    angular.module('officeAddin').controller('homeController', ['$scope', '$document', 'dataService', "proxyPageUrl", "sharePointUrl", homeController]);

    /**
     * Controller constructor
     */
    function homeController($scope, $document, dataService, proxyPageUrl, sharePointUrl) {
        var vm = this;
        vm.searchQuery = '';
        vm.loaded = false;
        vm.loading = false;
        vm.currentDoc = {
            location: null,
            title: null
        };
        vm.boards = [];
        vm.search = searchQuery;
        vm.noBoardsShown = function () {
            if (vm.noBoardsLoaded()) return false;
            var noBoards = true;
            vm.boards.forEach(function (b, i) {
                if (b.controller.visibleInFilter) {
                    noBoards = false;
                }
            });
            return noBoards;
        };
        vm.noBoardsLoaded = function () {
            return vm.loaded && vm.boards.length == 0;
        };

        activate();

        function activate() {
            initProxyPageContainer();
            loadDocumentLocation();                     
        }

        function initProxyPageContainer() {
            $("#proxyPageContainer").html("<iframe id='spProxyPage' src='" + sharePointUrl + proxyPageUrl + "' style='width: 1px; height: 1px;'></iframe>");
        }

        function searchQuery(clearSearch) {
            var query = vm.searchQuery.toLowerCase();
            if (query.length < 3 || clearSearch) {
                vm.boards.forEach(function (b, i) {
                    b.controller.visibleInFilter = true;
                });
            } else {
                vm.boards.forEach(function (b, i) {
                    b.controller.visibleInFilter = b.title.toLowerCase().indexOf(query) >= 0;
                });
            }
            $scope.$applyAsync();
        }

        function loadDocumentLocation() {
            //Note: This will return "undefined" when the document is embedded in a webpage.
            Office.context.document.getFilePropertiesAsync(
              function (asyncResult) {
                  if (asyncResult.status == "failed") {
                      //TODO: later.
                      //showMessage("Action failed with error: " + asyncResult.error.message);
                  } else {
                      vm.currentDoc.location = asyncResult.value.url;
                      loadDocumentInfo();
                  }
              }
            );
        }

        function loadDocumentInfo() {
            dataService.getDocumentByUrl(vm.currentDoc.location).then(function (document) {
                vm.currentDoc.title = document.title;
                loadAllBoards();
            });
        }

        function loadAllBoards() {
            vm.loading = true;
            vm.boards.length = 0;

            dataService.getAllBoards().then(function (boards) {
                boards.forEach(function (board) {
                    vm.boards.push(board);
                });
                vm.loading = false;
                vm.loaded = true;
            });
        }
    }    
})();