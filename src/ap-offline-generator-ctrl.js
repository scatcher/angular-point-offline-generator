'use strict';

angular.module('angularPoint')
  .controller('generateOfflineCtrl', function ($scope, $q, apDataService, apConfig, apDebugService, toastr) {

    /** Supported query types */
    $scope.operations = ['GetListItemChangesSinceToken', 'GetListItems'];

    /** State variables that are exposed to the view */
    $scope.state = {
      availableListFields: [],
      fileName: '',
      itemLimit: 0,
      operation: $scope.operations[0],
      query: '',
      selectedList: '',
      selectedListFields: [],
      siteUrl: apConfig.defaultUrl,
      xmlResponse: ''
    };

    $scope.listCollection = [];

    /** Get list/library definitions for the site */
    $scope.getLists = function () {
      apDataService.getCollection({
        operation: "GetListCollection",
        webURL: $scope.state.siteUrl
      }).then(function (dataArray) {
        $scope.listCollection.push.apply($scope.listCollection, dataArray);
        toastr.info($scope.listCollection.length + ' lists/libraries identified.');
      });
    };

    /** Make initial call */
    $scope.getLists();

    /** Main call to query list/library and capture XML response */
    $scope.getXML = function () {

      /** Empty out any previous values */
      $scope.listCollection.length = 0;
      $scope.state.xmlResponse = '';

      var payload = {
        operation: $scope.state.operation,
        listName: $scope.state.selectedList.ID,
        CAMLRowLimit: $scope.state.itemLimit,
        offlineXML: apConfig.offlineXML + $scope.state.fileName,
        webURL: $scope.state.siteUrl
      };

      /** Build custom CAMLViewFields if anything is identified in "Selected Fields" */
      if ($scope.state.selectedListFields.length > 0) {
        payload.CAMLViewFields = "<ViewFields>";
        _.each($scope.state.selectedListFields, function (fieldName) {
          payload.CAMLViewFields += "<FieldRef Name='" + fieldName + "' />";
        });
        payload.CAMLViewFields += "</ViewFields>";
      }

      /** Add query to payload if it's supplied */
      if ($scope.state.query.length > 0) {
        payload.CAMLQuery = $scope.state.query;
      }

      /** Use service wrapper to make the query */
      apDataService.serviceWrapper(payload)
        .then(function (response) {
          /** Update the visible XML response and allow for downloading */
          $scope.state.xmlResponse = apConfig.offline ?
            /** Get the string representation of the XML */
            (new XMLSerializer()).serializeToString(response) : response.responseText;
        });
    };

    /** Automatically highlight all text when the textarea is selected to
     * allow for easier copy/paste. */
    $scope.onTextClick = function ($event) {
      $event.target.select();
    };

    /** Fetch the available fields for the selected list */
    $scope.lookupListFields = function () {
      console.log("Looking up List Fields");
      $scope.state.availableListFields.length = 0;
      $scope.state.selectedListFields.length = 0;
      if (_.isObject($scope.state.selectedList) && $scope.state.selectedList.Title) {

        apDataService.getListFields({
          listName: $scope.state.selectedList.Name
        }).then(function (dataArray) {
          $scope.state.availableListFields.push.apply(
            $scope.state.availableListFields, dataArray
          );
          toastr.info($scope.state.availableListFields.length + " fields found.")
        });
      } else {
        toastr.error('Please ensure a list is selected before preceding.')
      }
    };

    /** Save the XML response to the local machine */
    $scope.saveXML = function () {
      apDebugService.saveXML($scope.state.xmlResponse, $scope.state.fileName);
    };

    /** Get list fields whenever the selected list changes */
    $scope.$watch('state.selectedList', function (newVal, oldVal) {
      if (!newVal) return;

      /** Remove spaces in name to take a guess at the name of the XML file */
      $scope.state.fileName = $scope.state.selectedList.Title.replace(' ','') + '.xml';

      $scope.lookupListFields();
      console.log(newVal);
    });
  });