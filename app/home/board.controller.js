(function () {
    'use strict';

    angular.module('officeAddin').controller('boardController', ['$scope', '$document', 'dataService', 'azureOrigin', boardController]);

    /**
     * Controller constructor
     */
    function boardController($scope, $document, dataService, azureOrigin) {
        var vm = this;
        vm.selected = false;
        vm.loading = false;
        vm.loaded = false;
        vm.docLocation = "";
        vm.getBoardDocuments = getBoardDocuments;
        vm.noDocs = function () { return $scope.boardDocuments.length == 0; }
        vm.$scope = $scope;

        $scope.boardDocuments = [];
        $scope.toggleBoardSelected = toggleBoardSelected;
        $scope.expandCollapsBoard = expandCollapsBoard;
        $scope.loadBoardDocuments = loadBoardDocuments;
        $scope.visibleInFilter = true;
        $scope.isOpened = false;
        $scope.getBoardClass = function () {
            var cl = "board ";
            if (vm.selected) {
                cl += "selected ";
            }
            if ($scope.isOpened) {
                cl += "opened ";
            }
            return cl;
        }
        activate();

        function activate() {
            //put here any initializers
        }        

        function loadBoardDocuments(currentDocLocation) {
            $scope.board.controller = this;
            vm.docLocation = currentDocLocation;
            getBoardDocuments($scope.board);
        }

        function expandCollapsBoard(event) {
            expandCollapsBoardUI(event);
            //getBoardDocuments(document);
        }

        function toggleBoardSelected(event) {
            event.originalEvent.preventDefault();
            //var elem = event.currentTarget;
            //$(elem).parents(".board").toggleClass('selected');
            vm.selected = !vm.selected;
            $scope.$applyAsync();

            //ToDo: send request to add/remove current document to/from board
            var message = { board: $scope.board.title, docLocation: vm.docLocation };
            //$document[0].getElementById("spProxy").contentWindow.postMessage(JSON.stringify(message), azureOrigin);
        }

        function expandCollapsBoardUI(event) {
            event.originalEvent.preventDefault();
            event.stopPropagation();

            var elem = event.currentTarget;
            var $board = $(elem).parents(".board");
            var files = $board.find('.board-files');
            $scope.isOpened ? files.hide('fast') : files.show('fast');

            $scope.isOpened = !$scope.isOpened;
            $scope.$applyAsync();

            return false;
        }

        //Gets all documents in current Delve Board
        function getBoardDocuments(board) {
            if (vm.loaded || vm.loading) return;
            vm.loading = true;
            $scope.boardDocuments.length = 0;

            dataService.getBoardDocuments(board).then(function (documents) {
                documents.forEach(function (document) {
                    $scope.boardDocuments.push(document);
                    if (document.url == vm.docLocation) {
                        vm.selected = true;
                    }
                });
                vm.loading = false;
                vm.loaded = true;
            });
        }
    }
})();