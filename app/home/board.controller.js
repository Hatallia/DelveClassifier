(function () {
  'use strict';

  angular.module('officeAddin')
    .controller('boardController', ['$scope', 'dataService', boardController]);

  /**
   * Controller constructor
   */
  function boardController($scope, dataService) {
    var vm = this;
    vm.selected = false;

    vm.loading = false;
    vm.loaded = false;
    $scope.boardDocuments = [];
    vm.currentBoard = null;
    
    
    vm.getBoardDocuments = getBoardDocuments;
    $scope.toggleBoardSelected = toggleBoardSelected;
    $scope.expandCollapsBoard = expandCollapsBoard;
    activate();

    function activate() {
      // if (Office.context.document) {
      //   Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, selectedTextChanged);
      // }
     // vm.getAllBoards();

    //  getDocumentLocation();
    }

    function expandCollapsBoard(event, document){
      expandCollapsBoardUI(event);
      getBoardDocuments(document);
    }

    function toggleBoardSelected(event){
      event.originalEvent.preventDefault();
      var elem = event.currentTarget;
      $(elem).parents(".board").toggleClass('selected');
      vm.selected = !vm.selected;
    }

    function expandCollapsBoardUI(event) {
        event.originalEvent.preventDefault();
        event.stopPropagation();
        var elem = event.currentTarget;
      
        var $board = $(elem).parents(".board");

        if ($board.hasClass('opened')) {
            $board.find(".board-icon .ms-Icon").removeClass('ms-Icon--caretDownRight').addClass('ms-Icon--caretRightOutline');
            $board.find('.board-files').hide('fast');
        } else {
            $board.find(".board-icon .ms-Icon").removeClass('ms-Icon--caretRightOutline').addClass('ms-Icon--caretDownRight');
            $board.find('.board-files').show('fast');
        }

        $board.toggleClass('opened');
        //$board.siblings('.board').find('.board-icon .ms-Icon').removeClass('ms-Icon--caretDownRight').addClass('ms-Icon--caretRightOutline');
        //$board.siblings('.board').removeClass('opened');
        //$board.siblings('.board').find('.board-files').hide('fast');
        return false;
    }

    function getBoardDocuments(board) {
      if (vm.loaded || vm.loading) return;
      vm.loading = true;
      $scope.boardDocuments.length = 0;

      dataService.getBoardDocuments(board).then(function (documents) {
        documents.forEach(function (document) {
          $scope.boardDocuments.push(document);
        });

        vm.loading = false;
        vm.loaded = true;
      });
    }
  }

})();