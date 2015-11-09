(function () {
  'use strict';

  angular.module('officeAddin')
    .controller('homeController', ['$scope', 'dataService', homeController]);

  /**
   * Controller constructor
   */
  function homeController($scope, dataService) {
    var vm = this;
    vm.searchQuery = '';
    // vm.searchQueryKeyDown = searchQueryKeyDown;
    vm.hasSearched = false;
    vm.loading = false;
    vm.documents = [];
    vm.getAllBoards = getAllBoards;
    activate();

    function activate() {
      // if (Office.context.document) {
      //   Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, selectedTextChanged);
      // }
      vm.getAllBoards();
    }
    
    // function selectedTextChanged() {
    //   Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
    //     function (result) {
    //       if (result.status === Office.AsyncResultStatus.Succeeded) {
    //         vm.searchQuery = result.value;
    //         $scope.$apply();
    //       }
    //       else {
    //         console.error(result.error.message);
    //       }
    //     });
    // }

    // function searchQueryKeyDown($event) {
    //   if ($event.keyCode === 13 &&
    //     vm.searchQuery.length > 0) {
    //     vm.findDocuments();
    //   }
    //   else {
    //     return true;
    //   }
    // }

    function getAllBoards() {
      vm.loading = true;
      vm.documents.length = 0;

      dataService.getAllBoards().then(function (documents) {
        documents.forEach(function (document) {
          vm.documents.push(document);
        });

        vm.loading = false;
        vm.hasSearched = true;
      });
    }
  }

})();