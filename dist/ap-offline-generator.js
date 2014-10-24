'use strict';

/**
 * @ngdoc function
 * @name angularPoint.generateOfflineCtrl
 * @description
 * Allows us to view lists within a site collection and download a XML representation of the list to use offline.  Allows
 * the user to set query parameters to filter the information that is returned.
 * @requires angularPoint.apDataService
 * @requires angularPoint.apConfig
 * @requires angularPoint.apExportService
 */
angular.module('angularPoint')
    .controller('generateOfflineCtrl', ["$scope", "$q", "apDataService", "apConfig", "apExportService", "toastr", function ($scope, $q, apDataService, apConfig, apExportService, toastr) {

        /** Supported query types */
        var operations = ['GetListItemChangesSinceToken', 'GetListItems'],
            queryOptions = [
                {attr: 'IncludeMandatoryColumns', val: false},
                {attr: 'IncludeAttachmentUrls', val: true},
                {attr: 'IncludeAttachmentVersion', val: false},
                {attr: 'ExpandUserField', val: false}
            ],
            /** State variables that are exposed to the view */
            state = {
                availableListFields: [],
                fileName: '',
                itemLimit: 0,
                operation: operations[0],
                query: '',
                selectedList: '',
                selectedListFields: [],
                siteUrl: apConfig.defaultUrl,
                xmlResponse: ''
            };

        $scope.listCollection = [];
        $scope.lookupListFields = lookupListFields;
        $scope.getLists = getLists;
        $scope.getXML = getXML;
        $scope.onTextClick = onTextClick;
        $scope.operations = operations;
        $scope.queryOptions = queryOptions;
        $scope.saveXML = saveXML;
        $scope.state = state;


        activate();


        /**===========PRIVATE===========*/

        function activate() {
            /** Make initial call */
            getLists();

            /** Get list fields whenever the selected list changes */
            $scope.$watch('state.selectedList', function (newVal, oldVal) {
                if (!newVal) return;

                /** Remove spaces in name to take a guess at the name of the XML file */
                state.fileName = state.selectedList.Title.replace(' ', '') + '.xml';

                $scope.lookupListFields();
                console.log(newVal);
            });

        }


        /** Get list/library definitions for the site */
        function getLists() {
            /** Ensure array is empty */
            $scope.listCollection.length = 0;
            apDataService.getCollection({
                operation: "GetListCollection",
                webURL: state.siteUrl
            }).then(function (dataArray) {
                $scope.listCollection.push.apply($scope.listCollection, dataArray);
                toastr.info($scope.listCollection.length + ' lists/libraries identified.');
            });
        }


        /**
         * @name stripQuery
         * @description
         * Allows us to paste in the query directly from a model without removing the apostrophe's or plus
         * signs used.
         * @param {string} str CAML query.
         * @returns {string} Cleaned up string.
         * @example
         * <h3>The following could be directly pasted into the CAML Query textarea</h3>
         * <pre>
         * '<Query>' +
         * '   <Where>' +
         * '       <Geq>' +
         * '           <FieldRef Name="To"/>' +
         * '           <Value Type="DateTime">' +
         * '               <Today OffsetDays="-30"/>' +
         * '           </Value>' +
         * '       </Geq>' +
         * '   </Where>' +
         * '   <OrderBy>' +
         * '       <FieldRef Name="From" Ascending="TRUE"/>' +
         * '   </OrderBy>' +
         * '</Query>'
         * </pre>
         */
        function stripQuery(str) {
            /** Remove apostrophe's */
            return str.replace(/'+/g, '')
            /** Remove plus signs */
                .replace(/\++/g, '');
        }

        /** Main call to query list/library and capture XML response */
        function getXML() {

            /** Empty out any previous values */
            state.xmlResponse = '';

            var payload = {
                operation: state.operation,
                listName: state.selectedList.ID,
                CAMLRowLimit: state.itemLimit,
                offlineXML: apConfig.offlineXML + state.fileName,
                CAMLQueryOptions: '',
                webURL: state.siteUrl
            };

            /** Build custom CAMLViewFields if anything is identified in "Selected Fields" */
            if (state.selectedListFields.length > 0) {
                payload.CAMLViewFields = "<ViewFields>";
                _.each(state.selectedListFields, function (fieldName) {
                    payload.CAMLViewFields += "<FieldRef Name='" + fieldName + "' />";
                });
                payload.CAMLViewFields += "</ViewFields>";
            }

            /** Add query to payload if it's supplied */
            var camlQuery = stripQuery(state.query);
            if (camlQuery.length > 0) {
                payload.CAMLQuery = camlQuery;
            }

            /** Add query options */
            payload.CAMLQueryOptions = buildQueryOptions(queryOptions);

            /** Use service wrapper to make the query */
            apDataService.serviceWrapper(payload)
                .then(function (response) {
                    /** Update the visible XML response and allow for downloading */
                    state.xmlResponse =
                    /** Get the string representation of the XML */
                        (new XMLSerializer()).serializeToString(response);
                });
        }

        /**
         * Generate something that looks like the following using settings from the view
         * @returns {string}
         * @example
         * <pre>
         * '<QueryOptions>' +
         * '   <IncludeMandatoryColumns>FALSE</IncludeMandatoryColumns>' +
         * '   <IncludeAttachmentUrls>TRUE</IncludeAttachmentUrls>' +
         * '   <IncludeAttachmentVersion>FALSE</IncludeAttachmentVersion>' +
         * '   <ExpandUserField>FALSE</ExpandUserField>' +
         * '</QueryOptions>',
         * </pre>
         */
        function buildQueryOptions(opts) {
            var xml = '<QueryOptions>';
            _.each(opts, function(option) {
                xml += '<' + option.attr + '>' + (option.val ? 'TRUE' : 'FALSE') + '</' + option.attr + '>';
            });
            xml += '</QueryOptions>';
            return xml;
        }

        /** Automatically highlight all text when the textarea is selected to
         * allow for easier copy/paste. */
        function onTextClick($event) {
            $event.target.select();
        }

        /** Fetch the available fields for the selected list */
        function lookupListFields() {
            console.log("Looking up List Fields");
            state.availableListFields.length = 0;
            state.selectedListFields.length = 0;
            if (_.isObject(state.selectedList) && state.selectedList.Title) {

                apDataService.getListFields({
                    listName: state.selectedList.Name
                }).then(function (dataArray) {
                    state.availableListFields.push.apply(
                        state.availableListFields, dataArray
                    );
                    toastr.info(state.availableListFields.length + " fields found.")
                });
            } else {
                toastr.error('Please ensure a list is selected before preceding.')
            }
        }

        /** Save the XML response to the local machine */
        function saveXML() {
            apExportService.saveXML(state.xmlResponse, state.fileName);
        }

    }]);
;angular.module('angularPoint').run(['$templateCache', function($templateCache) {
  'use strict';

  $templateCache.put('bower_components/angular-point-offline-generator/src/ap-offline-generator-view.html',
    "<div class=container><h3>Offline XML Generator</h3><p>Fill in the basic information for a list and make the request to SharePoint. The xml response will appear at the bottom of the page where you can then copy by Ctrl + A.</p><hr><form role=form><div class=form-group><label>Site URL</label><div class=input-group><input type=url class=form-control ng-model=state.siteUrl> <span class=input-group-btn><button title=\"Refresh list/libraries\" class=\"btn btn-success\" type=button ng-click=getLists()><i class=\"fa fa-refresh\"></i></button></span></div></div><div class=row><div class=col-xs-5><div class=form-group><label>List Name or GUID</label><select ng-model=state.selectedList ng-options=\"list.Title for list in listCollection\" class=form-control></select></div></div><div class=col-xs-3><div class=form-group><label>Number of Items to Return</label><input type=number class=form-control ng-model=state.itemLimit><p class=help-block>[0 = All Items]</p></div></div><div class=col-xs-4><div class=form-group><label>Operation</label><select class=form-control ng-model=state.operation ng-options=\"operation for operation in operations\"></select><p class=help-block>Operation to query with.</p></div></div></div><div class=row><div class=col-xs-12><div class=form-group><label>Selected Fields</label><div ui-select multiple ng-model=state.selectedListFields theme=bootstrap ng-disabled=\"listCollection.length < 1 || !state.selectedList\"><div ui-select-match placeholder=\"Select fields to include...\">{{$item.Name}}</div><div ui-select-choices repeat=\"field in state.availableListFields\">{{ field.Name }}</div></div><p class=help-block>Leaving this blank will return all fields visible in the default list view.</p></div></div></div><fieldset><legend>Query Options</legend><div class=row><div class=col-xs-3 ng-repeat=\"option in queryOptions\"><div class=form-group><label>{{ option.attr }}</label><input type=checkbox ng-model=option.val class=form-control></div></div></div></fieldset><div class=form-group><label>CAML Query (Optional)</label><textarea class=form-control ng-model=state.query rows=3></textarea></div><div class=row><div class=col-xs-6><div class=form-group><button type=submit class=\"btn btn-primary\" ng-click=getXML()>Make Request</button></div></div><div class=col-xs-6><fieldset ng-disabled=\"state.xmlResponse.length < 1\"><div class=form-group><div class=input-group><input class=form-control ng-model=state.fileName> <span class=input-group-btn><button class=\"btn btn-success\" ng-click=saveXML()>Save XML</button></span></div></div></fieldset></div></div><br><hr><h4>XML Response</h4><ol><li>Create a new offline file under the identified in apConfig.offlineXML (defaults to: \"app/dev\")</li><li>Use the same name as found in the model at \"model.list.title\" + .xml</li><li>Select the returned XML below by clicking in the textarea.</li><li>Add the XML to the newly created offline .xml file.</li></ol><div class=\"well well-sm\"><textarea name=xmlResponse class=form-control cols=30 rows=10 ng-model=state.xmlResponse ng-click=onTextClick($event)></textarea></div></form></div>"
  );

}]);
