/// <reference path="../typings/tsd.d.ts" />
var ap;
(function (ap) {
    var offlineGenerator;
    (function (offlineGenerator) {
        'use strict';
        /**
         * @ngdoc function
         * @name angularPoint.generateOfflineCtrl
         * @description
         * Allows us to view lists within a site collection and download a XML representation of the list to use offline.  Allows
         * the user to set query parameters to filter the information that is returned.
         * @requires angularPoint.apDataService
         * @requires angularPoint.apExportService
         */
        var OfflineGeneratorController = (function () {
            function OfflineGeneratorController($scope, $q, apDataService, apExportService, toastr) {
                this.apDataService = apDataService;
                this.apExportService = apExportService;
                this.toastr = toastr;
                this.availableListFields = [];
                this.fileName = '';
                this.itemLimit = 0;
                this.listCollection = [];
                this.operations = ['GetListItemChangesSinceToken', 'GetListItems'];
                this.query = '';
                this.queryOptions = [
                    { attr: 'IncludeMandatoryColumns', val: false },
                    { attr: 'IncludeAttachmentUrls', val: true },
                    { attr: 'IncludeAttachmentVersion', val: false },
                    { attr: 'ExpandUserField', val: false }
                ];
                this.selectedListFields = [];
                var vm = this;
                vm.operation = vm.operations[0];
                vm.siteUrl = $scope.siteUrl;
                vm.getLists();
                /** Get list fields whenever the selected list changes */
                $scope.$watch('vm.selectedList', function (newVal, oldVal) {
                    if (!newVal)
                        return;
                    /** Remove spaces in name to take a guess at the name of the XML file */
                    vm.fileName = vm.selectedList.Title.replace(' ', '') + '.xml';
                    vm.lookupListFields();
                    console.log(newVal);
                });
            }
            /** Get list/library definitions for the site */
            OfflineGeneratorController.prototype.getLists = function () {
                var vm = this;
                /** Ensure array is empty */
                vm.listCollection.length = 0;
                vm.apDataService.getCollection({
                    filterNode: 'List',
                    operation: "GetListCollection",
                    webURL: vm.siteUrl
                }).then(function (dataArray) {
                    vm.listCollection.push.apply(vm.listCollection, dataArray);
                    vm.toastr.info(vm.listCollection.length + ' lists/libraries identified.');
                });
            };
            /** Main call to query list/library and capture XML response */
            OfflineGeneratorController.prototype.getXML = function () {
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
                    _.each(vm.selectedListFields, function (fieldName) {
                        payload.CAMLViewFields += "<FieldRef Name=\"" + fieldName + "\" />";
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
                    .then(function (response) {
                    /** Update the visible XML response and allow for downloading */
                    vm.xmlResponse =
                        /** Get the string representation of the XML */
                        (new XMLSerializer()).serializeToString(response);
                });
            };
            /** Fetch the available fields for the selected list */
            OfflineGeneratorController.prototype.lookupListFields = function () {
                var vm = this;
                vm.availableListFields.length = 0;
                vm.selectedListFields.length = 0;
                if (_.isObject(vm.selectedList) && vm.selectedList.Title) {
                    vm.apDataService.getListFields({
                        listName: vm.selectedList.Name
                    }).then(function (dataArray) {
                        vm.availableListFields.push.apply(vm.availableListFields, dataArray);
                        vm.toastr.info(vm.availableListFields.length + " fields found.");
                    });
                }
                else {
                    vm.toastr.error('Please ensure a list is selected before preceding.');
                }
            };
            /** Automatically highlight all text when the textarea is selected to
            * allow for easier copy/paste. */
            OfflineGeneratorController.prototype.onTextClick = function ($event) {
                $event.target.select();
            };
            /** Save the XML response to the local machine */
            OfflineGeneratorController.prototype.saveXML = function () {
                var vm = this;
                vm.apExportService.saveXML(vm.xmlResponse, vm.fileName);
            };
            return OfflineGeneratorController;
        })();
        offlineGenerator.OfflineGeneratorController = OfflineGeneratorController;
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
            _.each(opts, function (option) {
                xml += "<" + option.attr + ">" + (option.val ? 'TRUE' : 'FALSE') + "<" + option.attr + ">";
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
                .replace(/\++/g, '');
        }
    })(offlineGenerator = ap.offlineGenerator || (ap.offlineGenerator = {}));
})(ap || (ap = {}));

/// <reference path="../typings/tsd.d.ts" />
/// <reference path="apOfflineGeneratorController.ts" />
var ap;
(function (ap) {
    var offlineGenerator;
    (function (offlineGenerator) {
        'use strict';
        /**
         * @example
         *    //Offline route
         *    .state('offline', {
         *        url: '/offline',
         *        template: `<ap-offline-generator site-url="${apConfig.defaultUrl}"></ap-offline-generator>`
        })
         */
        offlineGenerator.OfflineGeneratorDirective = function () {
            var directive = {
                controller: offlineGenerator.OfflineGeneratorController,
                controllerAs: 'vm',
                scope: {
                    siteUrl: '@'
                },
                templateUrl: 'apOfflineGeneratorTemplate.html'
            };
            return directive;
        };
    })(offlineGenerator = ap.offlineGenerator || (ap.offlineGenerator = {}));
})(ap || (ap = {}));

/// <reference path="../typings/tsd.d.ts" />
/// <reference path="apOfflineGeneratorDirective.ts" />
var ap;
(function (ap) {
    var offlineGenerator;
    (function (offlineGenerator) {
        'use strict';
        angular
            .module('angularPoint')
            .directive('apOfflineGenerator', offlineGenerator.OfflineGeneratorDirective);
    })(offlineGenerator = ap.offlineGenerator || (ap.offlineGenerator = {}));
})(ap || (ap = {}));

//# sourceMappingURL=angular-point-offline-generator.js.map
angular.module("angularPoint").run(["$templateCache", function($templateCache) {$templateCache.put("apOfflineGeneratorTemplate.html","<div class=\"container\">\n    <h3>Offline XML Generator</h3>\n\n    <p>Fill in the basic information for a list and make the request to SharePoint. The xml\n        response will appear at the bottom of the page where you can then copy by Ctrl + A.</p>\n    <hr/>\n    <form role=\"form\">\n        <div class=\"form-group\">\n            <label>Site URL</label>\n\n            <div class=\"input-group\">\n                <input type=\"url\" class=\"form-control\" ng-model=\"vm.siteUrl\">\n                  <span class=\"input-group-btn\">\n                    <button title=\"Refresh list/libraries\" class=\"btn btn-success\"\n                            type=\"button\" ng-click=\"vm.getLists()\">\n                        <i class=\"fa fa-refresh\"></i></button>\n                  </span>\n            </div>\n        </div>\n        <div class=\"row\">\n            <div class=\"col-xs-5\">\n                <div class=\"form-group\">\n                    <label>List Name or GUID</label>\n                    <select ng-model=\"vm.selectedList\"\n                            ng-options=\"list.Title for list in vm.listCollection\"\n                            class=\"form-control\"></select>\n                </div>\n            </div>\n            <div class=\"col-xs-3\">\n                <div class=\"form-group\">\n                    <label>Number of Items to Return</label>\n                    <input type=\"number\" class=\"form-control\" ng-model=\"vm.itemLimit\"/>\n\n                    <p class=\"help-block\">[0 = All Items]</p>\n                </div>\n            </div>\n            <div class=\"col-xs-4\">\n                <div class=\"form-group\">\n                    <label>Operation</label>\n                    <select class=\"form-control\" ng-model=\"vm.operation\"\n                            ng-options=\"operation for operation in vm.operations\"></select>\n\n                    <p class=\"help-block\">Operation to query with.</p>\n                </div>\n            </div>\n        </div>\n        <div class=\"row\">\n            <div class=\"col-xs-12\">\n                <div class=\"form-group\">\n                    <label>Selected Fields</label>\n\n                    <div ui-select multiple ng-model=\"vm.selectedListFields\" theme=\"bootstrap\"\n                               ng-disabled=\"listCollection.length < 1 || !vm.selectedList\">\n                        <div ui-select-match placeholder=\"Select fields to include...\">{{$item.Name}}</div>\n                        <div ui-select-choices repeat=\"field in vm.availableListFields\">\n                            {{ field.Name }}\n                        </div>\n                    </div>\n                    <p class=\"help-block\">Leaving this blank will return all fields visible in the default list\n                        view.</p>\n                </div>\n            </div>\n        </div>\n        <fieldset><legend>Query Options</legend>\n            <div class=\"row\">\n                <div class=\"col-xs-3\" ng-repeat=\"option in vm.queryOptions\">\n                    <div class=\"form-group\">\n                        <label>{{ option.attr }}</label>\n                        <input type=\"checkbox\" ng-model=\"option.val\" class=\"form-control\">\n                    </div>\n                </div>\n            </div>\n        </fieldset>\n        <div class=\"form-group\">\n            <label>CAML Query (Optional)</label>\n            <textarea class=\"form-control\" ng-model=\"vm.query\" rows=\"3\"></textarea>\n        </div>\n        <div class=\"row\">\n            <div class=\"col-xs-6\">\n                <div class=\"form-group\">\n                    <button type=\"submit\" class=\"btn btn-primary\" ng-click=\"vm.getXML()\">Make Request</button>\n                </div>\n            </div>\n            <div class=\"col-xs-6\">\n                <fieldset ng-disabled=\"vm.xmlResponse.length < 1\">\n                    <div class=\"form-group\">\n                        <div class=\"input-group\">\n                            <input type=\"text\" class=\"form-control\" ng-model=\"vm.fileName\">\n                           <span class=\"input-group-btn\">\n                              <button class=\"btn btn-success\"\n                                      ng-click=\"vm.saveXML()\">Save XML\n                              </button>\n                           </span>\n                        </div>\n                    </div>\n                </fieldset>\n            </div>\n        </div>\n        <br/>\n        <hr/>\n        <h4>XML Response</h4>\n        <ol>\n            <li>Create a new offline file under the identified in apConfig.offlineXML (defaults to: \"app/dev\")</li>\n            <li>Use the same name as found in the model at \"model.list.title\" + .xml</li>\n            <li>Select the returned XML below by clicking in the textarea.</li>\n            <li>Add the XML to the newly created offline .xml file.</li>\n        </ol>\n        <div class=\"well well-sm\">\n            <textarea name=\"xmlResponse\" class=\"form-control\" cols=\"30\" rows=\"10\"\n                      ng-model=\"vm.xmlResponse\" ng-click=\"vm.onTextClick($event)\"></textarea>\n        </div>\n    </form>\n</div>\n");}]);