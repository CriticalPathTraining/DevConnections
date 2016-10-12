'use strict';

var app = angular.module('AzureAdSpo');

app.controller('groupsController', ['$scope', 'azureGraphApiService', 'adalAuthenticationService', groupsController]);

function groupsController($scope, unifiedApiService) {

  unifiedApiService.getGroups().success(function (data) {
    $scope.groups = data.value;
  }).
  error(function (data, status, headers, config) {
    alert("Error getting groups");
  });


}