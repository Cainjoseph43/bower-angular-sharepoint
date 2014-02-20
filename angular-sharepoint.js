/**
 * @ngdoc overview
 * @name ExpertsInside.SharePoint
 *
 * @description
 * The main module which holds everything together.
 */
angular.module('ExpertsInside.SharePoint', ['ng']);
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
 * the need to interact with the low level $http service.
 *
 * @return {Object} A list "class" object with the default set of resource actions
 */
angular.module('ExpertsInside.SharePoint').factory('$spList', [
  '$spPageContextInfo',
  '$spRequestDigest',
  '$http',
  '$log',
  function ($spPageContextInfo, $spRequestDigest, $http, $log) {
    'use strict';
    var $spListMinErr = angular.$$minErr('$spList');
    var validParamKeys = [
        '$select',
        '$filter',
        '$orderby',
        '$top',
        '$skip',
        '$expand',
        '$sort'
      ];
    function List(name, defaults) {
      if (!name) {
        throw $spListMinErr('badargs', 'name cannot be blank.');
      }
      this.name = name.toString();
      var upcaseName = this.name.charAt(0).toUpperCase() + this.name.slice(1);
      this.defaults = angular.extend({ itemType: 'SP.Data.' + upcaseName + 'ListItem' }, defaults);
    }
    List.prototype = {
      $baseUrl: function () {
        return $spPageContextInfo.webServerRelativeUrl + '/_api/web/lists/getByTitle(\'' + this.name + '\')';
      },
      $buildHttpConfig: function (action, params, args) {
        var baseUrl = this.$baseUrl();
        var httpConfig = {
            url: baseUrl,
            params: params,
            headers: { accept: 'application/json;odata=verbose' },
            transformResponse: function (data) {
              var response = JSON.parse(data);
              if (angular.isDefined(response.d)) {
                response = response.d;
              }
              if (angular.isDefined(response.results)) {
                response = response.results;
              }
              return response;
            }
          };
        switch (action) {
        case 'get':
          httpConfig.url = baseUrl + '/items(' + args + ')';
          httpConfig.method = 'GET';
          break;
        case 'query':
          httpConfig.url = baseUrl + '/items';
          httpConfig.method = 'GET';
          break;
        case 'create':
          httpConfig.url = baseUrl + '/items';
          httpConfig.method = 'POST';
          angular.extend(httpConfig.headers, {
            'X-RequestDigest': $spRequestDigest(),
            'content-type': 'application/json;odata=verbose'
          });
          break;
        case 'save':
          httpConfig.url = args.__metdata.uri;
          httpConfig.method = 'POST';
          angular.extend(httpConfig.headers, {
            'X-HTTP-Method': 'MERGE',
            'X-RequestDigest': $spRequestDigest(),
            'IF-MATCH': args.__metadata.etag || '*',
            'content-type': 'application/json;odata=verbose'
          });
          break;
        }
        return httpConfig;
      },
      $normalizeParams: function (params) {
        params = angular.extend({}, params);
        //make a copy
        if (angular.isDefined(params)) {
          angular.forEach(params, function (value, key) {
            if (key.indexOf('$') !== 0) {
              delete params[key];
              key = '$' + key;
              params[key] = value;
            }
            if (angular.isArray(value)) {
              params[key] = value.join(',');
            }
            if (validParamKeys.indexOf(key) === -1) {
              $log.warn('Invalid param key: ' + key);
              delete params[key];
            }
          });
        }
        // cannot use angular.equals(params, {}) to check for empty object,
        // because angular.equals ignores properties prefixed with $
        if (params === null || JSON.stringify(params) === '{}') {
          params = undefined;
        }
        return params;
      },
      $createResult: function (emptyObject, httpConfig) {
        var result = emptyObject;
        result.$promise = $http(httpConfig).success(function (data) {
          angular.extend(result, data);
          return result;
        });
        return result;
      },
      get: function (id, params) {
        if (angular.isUndefined(id)) {
          throw $spListMinErr('badargs', 'id is required.');
        }
        params = this.$normalizeParams(params);
        var httpConfig = this.$buildHttpConfig('get', params, id);
        return this.$createResult({ Id: id }, httpConfig);
      },
      query: function (params) {
        params = this.$normalizeParams(params);
        var httpConfig = this.$buildHttpConfig('query', params);
        return this.$createResult([], httpConfig);
      },
      create: function (data) {
        var type = this.defaults.itemType;
        if (!type) {
          throw $spListMinErr('badargs', 'Cannot create an item without a valid type.' + 'Please set the default item type on the list (list.defaults.itemType).');
        }
        var itemDefaults = { __metadata: { type: type } };
        var item = angular.extend({}, itemDefaults, data);
        var httpConfig = this.$buildHttpConfig('create', undefined, item);
        return this.$createResult(item, httpConfig);
      },
      save: function (item) {
        if (angular.isUndefined(item.__metadata)) {
          throw 'o';
        }
        var httpConfig = this.$buildHttpConfig('save', undefined, item);
        return this.$createResult([], httpConfig);
      }
    };
    function listFactory(name, defaults) {
      return new List(name, defaults);
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
/**
 * @ngdoc Service
 * @name ExpertsInside.SharePoint.$spRequestDigest
 * @requires $window
 *
 * @description
 * Reads the request digest from the current page
 *
 * @return {String} request digest
 */
angular.module('ExpertsInside.SharePoint').factory('$spRequestDigest', [
  '$window',
  function ($window) {
    'use strict';
    var $spRequestDigestMinErr = angular.$$minErr('$spRequestDigest');
    return function () {
      var requestDigest = $window.document.getElementById('__REQUESTDIGEST');
      if (angular.isUndefined(requestDigest)) {
        throw $spRequestDigestMinErr('notfound', 'Cannot read request digest from DOM.');
      }
      return requestDigest.value;
    };
  }
]);