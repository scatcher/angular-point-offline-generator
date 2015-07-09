/// <reference path="../typings/tsd.d.ts" />

module ap.offlineGenerator {
    'use strict';

    interface IControllerScope extends ng.IScope {
        siteUrl: string;
    }

    /**
     * @ngdoc function
     * @name angularPoint.generateOfflineCtrl
     * @description
     * Allows us to view lists within a site collection and download a XML representation of the list to use offline.  Allows
     * the user to set query parameters to filter the information that is returned.
     * @requires angularPoint.apDataService
     * @requires angularPoint.apExportService
     */

    export class OfflineGeneratorController {
        availableListFields: IXMLFieldDefinition[] = [];
        fileName = '';
        itemLimit = 0;
        listCollection: IXMLList[] = [];
        operation: string;
        operations = ['GetListItemChangesSinceToken', 'GetListItems'];
        query = '';
        queryOptions = [
            { attr: 'IncludeMandatoryColumns', val: false },
            { attr: 'IncludeAttachmentUrls', val: true },
            { attr: 'IncludeAttachmentVersion', val: false },
            { attr: 'ExpandUserField', val: false }
        ]
        selectedList: IXMLList;
        selectedListFields: IXMLFieldDefinition[] = [];
        siteUrl: string;
        xmlResponse: string;
        constructor($scope: IControllerScope, $q: ng.IQService, private apDataService: DataService,
            private apExportService: ExportService, private toastr) {

            var vm = this;
            vm.operation = vm.operations[0];
            vm.siteUrl = $scope.siteUrl;


            vm.getLists();

            /** Get list fields whenever the selected list changes */
            $scope.$watch('vm.selectedList', function(newVal, oldVal) {
                if (!newVal) return;

                /** Remove spaces in name to take a guess at the name of the XML file */
                vm.fileName = vm.selectedList.Title.replace(' ', '') + '.xml';

                vm.lookupListFields();
                console.log(newVal);
            });


        }
        
        /** Get list/library definitions for the site */
        getLists() {
            var vm = this;
            /** Ensure array is empty */
            vm.listCollection.length = 0;
            vm.apDataService.getCollection({
                filterNode: 'List',
                operation: "GetListCollection",
                webURL: vm.siteUrl
            }).then(function(dataArray) {
                vm.listCollection.push.apply(vm.listCollection, dataArray);
                vm.toastr.info(vm.listCollection.length + ' lists/libraries identified.');
            });
        }
        
        /** Main call to query list/library and capture XML response */
        getXML() {
            var vm = this;
            /** Empty out any previous values */
            vm.xmlResponse = undefined;

            var payload = {
                operation: vm.operation,
                listName: vm.selectedList.ID,
                CAMLRowLimit: vm.itemLimit,
                // offlineXML: apConfig.offlineXML + vm.fileName,
                CAMLQuery: undefined,
                CAMLQueryOptions: '',
                CAMLViewFields: undefined,
                webURL: vm.siteUrl
            };

            /** Build custom CAMLViewFields if anything is identified in "Selected Fields" */
            if (vm.selectedListFields.length > 0) {
                payload.CAMLViewFields = "<ViewFields>";
                _.each(vm.selectedListFields, function(fieldName) {
                    payload.CAMLViewFields += `<FieldRef Name="${fieldName}" />`;
                });
                payload.CAMLViewFields += "</ViewFields>";
            }

            /** Add query to payload if it's supplied */
            var camlQuery = stripQuery(vm.query);
            if (camlQuery.length > 0) {
                payload.CAMLQuery = camlQuery;
            }

            /** Add query options */
            payload.CAMLQueryOptions = buildQueryOptions(vm.queryOptions);

            /** Use service wrapper to make the query */
            vm.apDataService.serviceWrapper(payload)
                .then(function(response) {
                    /** Update the visible XML response and allow for downloading */
                    vm.xmlResponse =
                    /** Get the string representation of the XML */
                    (new XMLSerializer()).serializeToString(response);
                });
        }
        
        /** Fetch the available fields for the selected list */
        lookupListFields() {
            var vm = this;
            vm.availableListFields.length = 0;
            vm.selectedListFields.length = 0;
            if (_.isObject(vm.selectedList) && vm.selectedList.Title) {

                vm.apDataService.getListFields({
                    listName: vm.selectedList.Name
                }).then(function(dataArray) {
                    vm.availableListFields.push.apply(vm.availableListFields, dataArray);
                    vm.toastr.info(vm.availableListFields.length + " fields found.")
                });
            } else {
                vm.toastr.error('Please ensure a list is selected before preceding.')
            }
        }
        
        /** Automatically highlight all text when the textarea is selected to
        * allow for easier copy/paste. */
        onTextClick($event) {
            $event.target.select();
        }
        
        /** Save the XML response to the local machine */
        saveXML() {
            var vm = this;
            vm.apExportService.saveXML(vm.xmlResponse, vm.fileName);
        }

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
            xml += `<${option.attr}>${option.val ? 'TRUE' : 'FALSE'}<${option.attr}>`;
        });
        xml += '</QueryOptions>';
        return xml;
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

}