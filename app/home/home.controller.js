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
    vm.searchQueryKeyDown = searchQueryKeyDown;
    vm.hasSearched = false;
    vm.loading = false;
    vm.documents = [];
    vm.docLocation = "";
    vm.getAllBoards = getAllBoards;
    vm.getFilteredBoards = getFilteredBoards;
    activate();

    function activate() {
      // if (Office.context.document) {
      //   Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, selectedTextChanged);
      // }
      vm.getAllBoards();

      getDocumentLocation();
    }

    function getDocumentLocation()
    {
      //Note: This will return "undefined" when the document is embedded in a webpage.
      Office.context.document.getFilePropertiesAsync(
        function (asyncResult) {
          if (asyncResult.status == "failed") {
            //TODO: later.
            //showMessage("Action failed with error: " + asyncResult.error.message);
          } else {
            vm.docLocation = asyncResult.value.url;
          }
        }
      );
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

    function searchQueryKeyDown($event) {
      if ($event.keyCode === 13 || 
        vm.searchQuery.length > 2) {
        
        vm.getFilteredBoards(vm.searchQuery);
      }
      else {
        return true;
      }
    }

    function getFilteredBoards(query) {
      vm.loading = true;
      vm.documents.length = 0;

      dataService.getFilteredBoards(query).then(function (documents) {
        documents.forEach(function (document) {
          vm.documents.push(document);
        });

        vm.loading = false;
        vm.hasSearched = true;
      });
    }

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