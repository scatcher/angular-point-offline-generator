'use strict';

/**
 * @ngdoc overview
 * @module
 * @name Angular Point Offline Generator
 * @description
 * A tool developed to assist with the angular-point efforts.
 */
angular.module('demo', [
  'ui.bootstrap',
  'ui.select',
  'ui.router',
  'toastr',

  //SP-Angular
  'angularPoint'

])
  .config(function ($stateProvider, $urlRouterProvider) {

    // For any unmatched url, redirect to /state1
    $urlRouterProvider.otherwise('/offline');

    // Now set up the states
    $stateProvider
      //Offline
      .state('offline', {
        url: '/offline',
        templateUrl: 'bower_components/angular-point-offline-generator/src/ap-offline-generator-view.html',
        controller: 'generateOfflineCtrl'
      });

  })
  .run(function (apConfig, toastrConfig) {
    console.log('Injector done loading all modules.');

    /** Set the default toast location */
    toastrConfig.positionClass = 'toast-bottom-right';

    /** Set the default site to look for lists, not really necessary if you're using list GUID's but is
     * helpful with the offline generator because it populates the list of available lists/libraries. */
    apConfig.defaultUrl = '//mysharepointserver.com/mysite';

     /** Set the folder where offline XML is stored */
    apConfig.offlineXML = 'offline-xml/';

  });

