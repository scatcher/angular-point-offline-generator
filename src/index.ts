/// <reference path="../typings/tsd.d.ts" />
/// <reference path="apOfflineGeneratorDirective.ts" />

module ap.offlineGenerator {
    'use strict';

    angular
        .module('angularPoint')
        .directive('apOfflineGenerator', OfflineGeneratorDirective);
}