'use strict';

// !!! BEFORE YOU RUN THIS SAMPLE APP !!!
// --------------------------------------

// update the GUID for the client_id variable 
// in the initializeADALSettings function
var client_id = "0a415fb6-20e8-4b9e-b333-a9ffba4e34f6";


(function () {

  var app = angular.module("AzureAdSpo", ['ngRoute', 'AdalAngular']);

  app.filter('unsafe', function ($sce) {
    return function (val) {
      return $sce.trustAsHtml(val);
    };
  });


  app.config(['$routeProvider', '$httpProvider', 'adalAuthenticationServiceProvider', initializeApp]);


  function initializeApp($routeProvider, $httpProvider, adalProvider) {

    initializeADALSettings($httpProvider, adalProvider);

    initializeRouteMap($routeProvider);



  }

  function initializeADALSettings($httpProvider, adalProvider) {

    var endpoints = {
      "https://graph.windows.net/": "https://graph.windows.net/",
    };

    var adalProviderSettings = {
      tenant: 'common',
      clientId: client_id,
      extraQueryParameter: 'nux=1',
      endpoints: endpoints,
      cacheLocation: 'localStorage' // enable this for IE, as sessionStorage does not work for localhost.
    };

    adalProvider.init(adalProviderSettings, $httpProvider);
   
  }

  function initializeRouteMap($routeProvider) {

    // config app's route map
    $routeProvider
      .when("/",
           { templateUrl: 'App/views/home.html', controller: "homeController" })
      .when("/users",
           { templateUrl: 'App/views/users.html', controller: "usersController", requireADLogin: true })
      .when("/groups",
           { templateUrl: 'App/views/groups.html', controller: "groupsController", requireADLogin: true })
      .when("/tenancy",
           { templateUrl: 'App/views/tenancy.html', controller: "tenancyController", requireADLogin: true })
      .when("/userClaims",
           { templateUrl: 'App/views/userClaims.html', controller: "userClaimsController", requireADLogin: true })
      .otherwise({ redirectTo: "/" });
  }

})();