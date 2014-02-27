'use strict';
/**
 * @ngdoc overview
 * @name ExpertsInside.SharePoint
 *
 * @description
 * The main module which holds everything together.
 */
angular.module('ExpertsInside.SharePoint', ['ng']).run(function () {
  var sharepointMinErr = angular.$$minErr('sharepoint');
  if (angular.isUndefined(ShareCoffee)) {
    throw sharepointMinErr('noShareCoffee', 'angular-sharepoint depends on ShareCoffee to do its job.' + 'Either include the bundled ShareCoffee + angular-sharepoint file ' + 'or include ShareCoffe seperately before angular-sharepoint.');
  }
});
/**
 * @ngdoc service
 * @name ExpertsInside.SharePoint.$spList
 * @requires $spRest
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
  '$spRest',
  '$http',
  function ($spRest, $http) {
    'use strict';
    var $spListMinErr = angular.$$minErr('$spList');
    function List(name, options) {
      if (!name) {
        throw $spListMinErr('badargs', 'name cannot be blank.');
      }
      if (!angular.isObject(options)) {
        options = {};
      }
      this.name = name.toString();
      var upcaseName = this.name.charAt(0).toUpperCase() + this.name.slice(1);
      this.settings = {
        itemType: 'SP.Data.' + upcaseName + 'ListItem',
        readOnlyFields: angular.extend([
          'Author',
          'Editor',
          'Created',
          'Modified'
        ], options.readOnlyFields)
      };
      this.queries = {};
    }
    List.prototype = {
      $baseUrl: function () {
        return 'web/lists/getByTitle(\'' + this.name + '\')';
      },
      $createPayload: function (data) {
        var payload = angular.extend({}, data);
        angular.forEach(this.settings.readOnlyFields, function (readOnlyField) {
          delete payload[readOnlyField];
        });
        return payload;
      },
      $buildHttpConfig: function (action, params, args) {
        var baseUrl = this.$baseUrl(), httpConfig;
        switch (action) {
        case 'get':
          httpConfig = ShareCoffee.REST.build.read.for.angularJS({ url: baseUrl + '/items(' + args + ')' });
          break;
        case 'query':
          httpConfig = ShareCoffee.REST.build.read.for.angularJS({ url: baseUrl + '/items' });
          break;
        case 'create':
          httpConfig = ShareCoffee.REST.build.create.for.angularJS({
            url: baseUrl + '/items',
            payload: angular.toJson(this.$createPayload(args))
          });
          break;
        case 'save':
          httpConfig = ShareCoffee.REST.build.update.for.angularJS({
            url: baseUrl,
            payload: angular.toJson(this.$createPayload(args))
          });
          httpConfig.url = args.__metadata.uri;
          // ShareCoffe doesnt work with absolute urls atm
          break;
        case 'delete':
          httpConfig = ShareCoffee.REST.build.delete.for.angularJS({ url: baseUrl });
          httpConfig.url = args.__metadata.uri;
          break;
        }
        httpConfig.url = $spRest.appendQueryString(httpConfig.url, params);
        httpConfig.transformResponse = $spRest.transformResponse;
        return httpConfig;
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
        var httpConfig = this.$buildHttpConfig('get', params, id);
        return this.$createResult({ Id: id }, httpConfig);
      },
      query: function (params) {
        var httpConfig = this.$buildHttpConfig('query', params);
        return this.$createResult([], httpConfig);
      },
      create: function (data) {
        var type = this.settings.itemType;
        if (!type) {
          throw $spListMinErr('badargs', 'Cannot create an item without a valid type.' + 'Please set the default item type on the list (list.settings.itemType).');
        }
        var itemDefaults = { __metadata: { type: type } };
        var item = angular.extend({}, itemDefaults, data);
        var httpConfig = this.$buildHttpConfig('create', undefined, item);
        return this.$createResult(item, httpConfig);
      },
      save: function (item) {
        if (angular.isUndefined(item.__metadata)) {
          throw $spListMinErr('badargs', 'Item must have __metadata property.');
        }
        var httpConfig = this.$buildHttpConfig('save', undefined, item);
        return this.$createResult(item, httpConfig);
      },
      delete: function (item) {
        if (angular.isUndefined(item.__metadata)) {
          throw $spListMinErr('badargs', 'Item must have __metadata property.');
        }
        var httpConfig = this.$buildHttpConfig('delete', undefined, item);
        return this.$createResult(item, httpConfig);
      },
      addNamedQuery: function (name, createParams) {
        var me = this;
        this.queries[name] = function () {
          var params = createParams.apply(me, arguments);
          return me.query(params);
        };
        return me;
      }
    };
    function listFactory(name, options) {
      return new List(name, options);
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
angular.module('ExpertsInside.SharePoint').factory('$spRest', [
  '$log',
  function ($log) {
    'use strict';
    var validParamKeys = [
        '$select',
        '$filter',
        '$orderby',
        '$top',
        '$skip',
        '$expand',
        '$sort'
      ];
    function getKeysSorted(obj) {
      var keys = [];
      if (angular.isUndefined(obj) || obj === null) {
        return keys;
      }
      for (var key in obj) {
        if (obj.hasOwnProperty(key)) {
          keys.push(key);
        }
      }
      return keys.sort();
    }
    var $spRest = {
        transformResponse: function (json) {
          var response = {};
          if (angular.isDefined(json) && json !== null && json !== '') {
            response = angular.fromJson(json);
          }
          if (angular.isObject(response) && angular.isDefined(response.d)) {
            response = response.d;
          }
          if (angular.isObject(response) && angular.isDefined(response.results)) {
            response = response.results;
          }
          return response;
        },
        buildQueryString: function (params) {
          var parts = [];
          var keys = getKeysSorted(params);
          angular.forEach(keys, function (key) {
            var value = params[key];
            if (value === null || angular.isUndefined(value)) {
              return;
            }
            if (angular.isArray(value)) {
              value = value.join(',');
            }
            if (angular.isObject(value)) {
              value = angular.toJson(value);
            }
            parts.push(key + '=' + value);
          });
          var queryString = parts.join('&');
          return queryString;
        },
        normalizeParams: function (params) {
          params = angular.extend({}, params);
          //make a copy
          if (angular.isDefined(params)) {
            angular.forEach(params, function (value, key) {
              if (key.indexOf('$') !== 0) {
                delete params[key];
                key = '$' + key;
                params[key] = value;
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
        appendQueryString: function (url, params) {
          params = $spRest.normalizeParams(params);
          var queryString = $spRest.buildQueryString(params);
          if (queryString !== '') {
            url += (url.indexOf('?') === -1 ? '?' : '&') + queryString;
          }
          return url;
        }
      };
    return $spRest;
  }
]);