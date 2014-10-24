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
    .controller('generateOfflineCtrl', function ($scope, $q, apDataService, apConfig, apExportService, toastr) {

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

    });
