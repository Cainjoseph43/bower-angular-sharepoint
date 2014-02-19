/**
 * @ngdoc overview
 * @name ExpertsInside.SharePoint
 *
 * @description
 * The main module which holds everything together.
 */
angular.module('ExpertsInside.SharePoint', ['ng']);
var $spListMinErr = angular.$$minErr('$spList');
/**
 * @ngdoc service
 * @name ExpertsInside.SharePoint.$spList
 * @requires $spPageContextInfo
 *
 * @description
 * A factory which creates a list object that lets you interact with SharePoint Lists via the
 * SharePoint REST API
 *
 * The returned list object has action methods which provide high-level behaviors without
 * the need to interact with the low level {@link ng.$http $http} service.
 *
 * @return {Object} A list "class" object with the default set of resource actions
 */
angular.module('ExpertsInside.SharePoint').factory('$spList', [
  '$spPageContextInfo',
  '$http',
  function ($spPageContextInfo, $http) {
    'use strict';
    function List(name, defaults) {
      this.name = name;
      this.defaults = defaults;
    }
    List.prototype = {
      $baseUrl: function () {
        return $spPageContextInfo.webServerRelativeUrl + '/_api/web/lists/getByTitle(\'' + this.name + '\')';
      },
      $buildHttpConfig: function (method, args) {
        var baseUrl = this.$baseUrl();
        var httpConfig = {
            method: method,
            url: baseUrl,
            headers: { accept: 'application/json;odata=verbose' },
            transformResponse: function (data) {
              var response = JSON.parse(data).d;
              if (angular.isDefined(response.results)) {
                response = response.results;
              }
              return response;
            }
          };
        switch (method) {
        case 'get':
          httpConfig.url = baseUrl + '/items(' + args.id + ')';
          break;
        }
        return httpConfig;
      },
      get: function (id, options) {
        if (angular.isUndefined(id)) {
          throw $spListMinErr('badargs', 'id is required.');
        }
        options = angular.extend({ id: id }, options);
        var httpConfig = this.$buildHttpConfig('get', options);
        return $http(httpConfig);
      }
    };
    function listFactory(name) {
      return new List(name);
    }
    listFactory.List = List;
    return listFactory;
  }
]);
/**
 * @ngdoc object
 * @name ExpertsInside.SharePoint.$spPageContextInfo
 * @requires $window, $rootScope
 *
 * @description
 * Wraps the global '_spPageContextInfo' object in an angular service
 *
 * @return {Object} $spPageContextInfo Copy of the global '_spPageContextInfo' object
 */
angular.module('ExpertsInside.SharePoint').factory('$spPageContextInfo', [
  '$rootScope',
  '$window',
  function ($rootScope, $window) {
    'use strict';
    var $spPageContextInfo = {};
    angular.copy($window._spPageContextInfo, $spPageContextInfo);
    $rootScope.$watch(function () {
      return $window._spPageContextInfo;
    }, function (spPageContextInfo) {
      angular.copy(spPageContextInfo, $spPageContextInfo);
    }, true);
    return $spPageContextInfo;
  }
]);