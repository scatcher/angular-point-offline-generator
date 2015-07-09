/// <reference path="../typings/tsd.d.ts" />
/// <reference path="apOfflineGeneratorController.ts" />

module ap.offlineGenerator {
    'use strict';

    /**
     * @example
     *    //Offline route
     *    .state('offline', {
     *        url: '/offline',
     *        template: `<ap-offline-generator site-url="${apConfig.defaultUrl}"></ap-offline-generator>`
    })
     */

    export var OfflineGeneratorDirective = () => {
        var directive = {
            controller: OfflineGeneratorController,
            controllerAs: 'vm',
            scope: {
                siteUrl: '@'
            },
            templateUrl: 'apOfflineGeneratorTemplate.html'
        };
        return directive;
    }

}
