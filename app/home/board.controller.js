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
        vm.docInfo = {
            location: "",
            title: ""
        };
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

        function loadBoardDocuments(currentDocInfo) {
            $scope.board.controller = this;
            vm.docInfo.location = currentDocInfo.location;
            vm.docInfo.title = currentDocInfo.title;
            getBoardDocuments($scope.board);
        }

        function expandCollapsBoard(event) {
            expandCollapsBoardUI(event);
            //getBoardDocuments(document);
			dataService.getFormDigest();
        }

        function toggleBoardSelected(event) {
            if (event.currentTarget.querySelector(".ms-Icon.ms-Icon--arrowUpRight") == event.target) {
                return;
            }
            event.originalEvent.preventDefault();

            //update list of documents for current board
            if (vm.selected) {
                $scope.boardDocuments.forEach(function (d, i) {
                    if (d.url == vm.docInfo.location) {
                        $scope.boardDocuments.splice(i, 1);
                    }
                });
            }
            else {
                $scope.boardDocuments.push({
                    url: vm.docInfo.location,
                    title: vm.docInfo.title
                });
            }
            sortBoardDocuments();
            
            //change state of board
            vm.selected = !vm.selected;
            $scope.$applyAsync();

            //ToDo: send request to add/remove current document to/from board
            var message = { board: $scope.board.title, docLocation: vm.docInfo.location };
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
                    if (document.url == vm.docInfo.location) {
                        vm.selected = true;
                    }
                    sortBoardDocuments();
                });
                vm.loading = false;
                vm.loaded = true;
            });
        }

        function sortBoardDocuments() {
            $scope.boardDocuments.sort(function (a, b) {
                if (a.title > b.title) {
                    return 1;
                }
                if (a.title < b.title) {
                    return -1;
                }
                // a must be equal to b
                return 0;
            });
        }
    }
})();