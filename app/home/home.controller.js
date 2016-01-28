(function () {
    'use strict';

    angular.module('officeAddin').controller('homeController', ['$scope', '$document', 'dataService', 'proxyHackUrl', homeController]);

    /**
     * Controller constructor
     */
    function homeController($scope, $document, dataService, proxyHackUrl) {
        var vm = this;
        vm.searchQuery = '';
        vm.loaded = false;
        vm.loading = false;
        vm.currentDocLocation = null;
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
            getDocumentLocation();
            loadAllBoards();            
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
            console.log("Clear Search: " + clearSearch);
            $scope.$applyAsync();
        }

        function getDocumentLocation() {
            //Note: This will return "undefined" when the document is embedded in a webpage.
            Office.context.document.getFilePropertiesAsync(
              function (asyncResult) {
                  if (asyncResult.status == "failed") {
                      //TODO: later.
                      //showMessage("Action failed with error: " + asyncResult.error.message);
                  } else {
                      vm.currentDocLocation = asyncResult.value.url;
                  }
              }
            );
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