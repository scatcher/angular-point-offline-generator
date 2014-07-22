'use strict';

angular.module('angularPoint')
  .controller('generateOfflineCtrl', ["$scope", "$q", "apDataService", "apConfig", "apDebugService", "toastr", function ($scope, $q, apDataService, apConfig, apDebugService, toastr) {

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
  }]);;angular.module('angularPoint').run(['$templateCache', function($templateCache) {
  'use strict';

  $templateCache.put('bower_components/angular-point-offline-generator/src/ap-offline-generator-view.html',
    "<div class=container><h3>Offline XML Generator</h3><p>Fill in the basic information for a list and make the request to SharePoint. The xml response will appear at the bottom of the page where you can then copy by Ctrl + A.</p><hr><form role=form><div class=form-group><label>Site URL</label><div class=input-group><input type=url class=form-control ng-model=state.siteUrl> <span class=input-group-btn><button title=\"Refresh list/libraries\" class=\"btn btn-success\" type=button ng-click=getLists()><i class=\"fa fa-refresh\"></i></button></span></div></div><div class=row><div class=col-xs-5><div class=form-group><label>List Name or GUID</label><select ng-model=state.selectedList ng-options=\"list.Title for list in listCollection\" class=form-control></select></div></div><div class=col-xs-3><div class=form-group><label>Number of Items to Return</label><input type=number class=form-control ng-model=state.itemLimit><p class=help-block>[0 = All Items]</p></div></div><div class=col-xs-4><div class=form-group><label>Operation</label><select class=form-control ng-model=state.operation ng-options=\"operation for operation in operations\"></select><p class=help-block>Operation to query with.</p></div></div></div><div class=row><div class=col-xs-12><div class=form-group><label>Selected Fields</label><select multiple ui-select2 ng-model=state.selectedListFields style=\"width: 100%\" ng-disabled=\"listCollection.length < 1 || !state.selectedList\"><option ng-repeat=\"field in state.availableListFields\" value={{field.Name}}>{{ field.Name }}</option></select><p class=help-block>Leaving this blank will return all fields visible in the default list view.</p></div></div></div><div class=form-group><label>CAML Query (Optional)</label><textarea class=form-control ng-model=state.query rows=3></textarea></div><div class=row><div class=col-xs-6><div class=form-group><button type=submit class=\"btn btn-primary\" ng-click=getXML()>Make Request</button></div></div><div class=col-xs-6><fieldset ng-disabled=\"state.xmlResponse.length < 1\"><div class=form-group><div class=input-group><input class=form-control ng-model=state.fileName> <span class=input-group-btn><button class=\"btn btn-success\" ng-click=saveXML()>Save XML</button></span></div></div></fieldset></div></div><br><hr><h4>XML Response</h4><ol><li>Create a new offline file under the identified in apConfig.offlineXML (defaults to: \"app/dev\")</li><li>Use the same name as found in the model at \"model.list.title\" + .xml</li><li>Select the returned XML below by clicking in the textarea.</li><li>Add the XML to the newly created offline .xml file.</li></ol><div class=\"well well-sm\"><textarea name=xmlResponse class=form-control cols=30 rows=10 ng-model=state.xmlResponse ng-click=onTextClick($event)></textarea></div></form></div>"
  );

}]);
