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
    function listFactory(name, options) {
      if (!angular.isString(name) || name === '') {
        throw $spListMinErr('badargs', 'name must be a nen-empty string.');
      }
      if (!angular.isObject(options)) {
        options = {};
      }
      var upcaseName = name.charAt(0).toUpperCase() + name.slice(1);
      function ListItem(data) {
        angular.extend(this, data);
      }
      ListItem.$$listName = name;
      ListItem.getListName = function () {
        return ListItem.$$listName;
      };
      ListItem.$$listRelativeUrl = 'web/lists/getByTitle(\'' + name + '\')';
      ListItem.$decorateResult = function (result, httpConfig) {
        if (!angular.isArray(result) && !(result instanceof ListItem)) {
          result = new ListItem(result);
        }
        if (angular.isUndefined(result.$resolved)) {
          result.$resolved = false;
        }
        result.$promise = $http(httpConfig).then(function (response) {
          var data = response.data;
          if (angular.isArray(result) && angular.isArray(data)) {
            angular.forEach(data, function (item) {
              result.push(new ListItem(item));
            });
          } else if (angular.isObject(result)) {
            if (angular.isArray(data)) {
              if (data.length === 1) {
                angular.extend(result, data[0]);
              } else {
                throw $spListMinErr('badresponse', 'Expected response to contain an array with one object but got {1}', data.length);
              }
            } else if (angular.isObject(data)) {
              angular.extend(result, data);
            }
          } else {
            throw $spListMinErr('badresponse', 'Expected response to contain an {0} but got an {1}', angular.isArray(result) ? 'array' : 'object', angular.isArray(data) ? 'array' : 'object');
          }
          var responseEtag;
          if (response.status === 204 && angular.isString(responseEtag = response.headers('ETag'))) {
            result.__metadata.etag = responseEtag;
          }
          result.$resolved = true;
          return result;
        });
        return result;
      };
      ListItem.get = function (id, query) {
        if (angular.isUndefined(id) || id === null) {
          throw $spListMinErr('badargs', 'id is required.');
        }
        var result = { Id: id };
        var httpConfig = $spRest.buildHttpConfig(ListItem.$$listRelativeUrl, 'get', {
            id: id,
            query: query
          });
        return ListItem.$decorateResult(result, httpConfig);
      };
      ListItem.query = function (query, options) {
        var result = angular.isDefined(options) && options.singleResult ? {} : [];
        var httpConfig = $spRest.buildHttpConfig(ListItem.$$listRelativeUrl, 'query', { query: angular.extend({}, ListItem.prototype.$settings.queryDefaults, query) });
        return ListItem.$decorateResult(result, httpConfig);
      };
      ListItem.create = function (item, query) {
        if (!(angular.isObject(item) && item instanceof ListItem)) {
          throw $spListMinErr('badargs', 'item must be a ListItem instance.');
        }
        var type = item.$settings.itemType;
        if (!type) {
          throw $spListMinErr('badargs', 'Cannot create an item without a valid type');
        }
        item.__metadata = { type: type };
        var httpConfig = $spRest.buildHttpConfig(ListItem.$$listRelativeUrl, 'create', {
            item: item,
            query: angular.extend({}, item.$settings.queryDefaults, query)
          });
        return ListItem.$decorateResult(item, httpConfig);
      };
      ListItem.update = function (item, options) {
        if (!(angular.isObject(item) && item instanceof ListItem)) {
          throw $spListMinErr('badargs', 'item must be a ListItem instance.');
        }
        options = angular.extend({}, options, { item: item });
        var httpConfig = $spRest.buildHttpConfig(ListItem.$$listRelativeUrl, 'update', options);
        return ListItem.$decorateResult(item, httpConfig);
      };
      ListItem.save = function (item, options) {
        if (angular.isDefined(item.__metadata) && angular.isDefined(item.__metadata.id)) {
          return this.update(item, options);
        } else {
          var query = angular.isObject(options) ? options.query : undefined;
          return this.create(item, query);
        }
      };
      ListItem.delete = function (item) {
        if (!(angular.isObject(item) && item instanceof ListItem)) {
          throw $spListMinErr('badargs', 'item must be a ListItem instance.');
        }
        var httpConfig = $spRest.buildHttpConfig(ListItem.$$listRelativeUrl, 'delete', { item: item });
        return ListItem.$decorateResult(item, httpConfig);
      };
      ListItem.queries = {};
      ListItem.addNamedQuery = function (name, createQuery, options) {
        ListItem.queries[name] = function () {
          var query = angular.extend({}, ListItem.prototype.$settings.queryDefaults, createQuery.apply(ListItem, arguments));
          return ListItem.query(query, options);
        };
        return ListItem;
      };
      ListItem.prototype = {
        $settings: {
          itemType: 'SP.Data.' + upcaseName + 'ListItem',
          readOnlyFields: angular.extend([
            'AttachmentFiles',
            'Attachments',
            'Author',
            'AuthorId',
            'ContentType',
            'ContentTypeId',
            'Created',
            'Editor',
            'EditorId',
            'FieldValuesAsHtml',
            'FieldValuesAsText',
            'FieldValuesForEdit',
            'File',
            'FileSystemObjectType',
            'FirstUniqueAncestorSecurableObject',
            'Folder',
            'GUID',
            'Modified',
            'OData__UIVersionString',
            'ParentList',
            'RoleAssignments'
          ], options.readOnlyFields),
          queryDefaults: angular.extend({}, options.queryDefaults)
        },
        $save: function (options) {
          return ListItem.save(this, options).$promise;
        },
        $delete: function () {
          return ListItem.delete(this).$promise;
        }
      };
      return ListItem;
    }
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
    var $spRestMinErr = angular.$$minErr('$spRest');
    var unique = function (arr) {
      return arr.reduce(function (r, x) {
        if (r.indexOf(x) < 0) {
          r.push(x);
        }
        return r;
      }, []);
    };
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
              value = unique(value).join(',');
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
          var queryString = $spRest.buildQueryString(params);
          if (queryString !== '') {
            url += (url.indexOf('?') === -1 ? '?' : '&') + queryString;
          }
          return url;
        },
        createPayload: function (item) {
          var payload = angular.extend({}, item);
          if (angular.isDefined(item.$settings) && angular.isDefined(item.$settings.readOnlyFields)) {
            angular.forEach(item.$settings.readOnlyFields, function (readOnlyField) {
              delete payload[readOnlyField];
            });
          }
          return angular.toJson(payload);
        },
        buildHttpConfig: function (listUrl, action, options) {
          var baseUrl = listUrl + '/items';
          var httpConfig = { url: baseUrl };
          action = angular.isString(action) ? action.toLowerCase() : '';
          options = angular.isDefined(options) ? options : {};
          var query = angular.isDefined(options.query) ? $spRest.normalizeParams(options.query) : {};
          switch (action) {
          case 'get':
            if (angular.isUndefined(options.id)) {
              throw $spRestMinErr('options:get', 'options must have an id');
            }
            httpConfig = ShareCoffee.REST.build.read.for.angularJS({ url: baseUrl + '(' + options.id + ')' });
            break;
          case 'query':
            httpConfig = ShareCoffee.REST.build.read.for.angularJS({ url: baseUrl });
            break;
          case 'create':
            if (angular.isUndefined(options.item)) {
              throw $spRestMinErr('options:create', 'options must have an item');
            }
            delete query.$expand;
            httpConfig = ShareCoffee.REST.build.create.for.angularJS({
              url: baseUrl,
              payload: $spRest.createPayload(options.item)
            });
            break;
          case 'update':
            if (angular.isUndefined(options.item)) {
              throw $spRestMinErr('options:update', 'options must have an item');
            }
            if (angular.isUndefined(options.item.__metadata)) {
              throw $spRestMinErr('options:update', 'options.item must have __metadata');
            }
            query = {};
            // does nothing or breaks things, so we ignore it
            var eTag = !options.force && angular.isDefined(options.item.__metadata) ? options.item.__metadata.etag : null;
            httpConfig = ShareCoffee.REST.build.update.for.angularJS({
              url: baseUrl,
              payload: $spRest.createPayload(options.item),
              eTag: eTag
            });
            httpConfig.url = options.item.__metadata.uri;
            // ShareCoffe doesnt work with absolute urls atm
            break;
          case 'delete':
            if (angular.isUndefined(options.item)) {
              throw $spRestMinErr('options:delete', 'options must have an item');
            }
            if (angular.isUndefined(options.item.__metadata)) {
              throw $spRestMinErr('options:delete', 'options.item must have __metadata');
            }
            httpConfig = ShareCoffee.REST.build.delete.for.angularJS({ url: baseUrl });
            httpConfig.url = options.item.__metadata.uri;
            // ShareCoffe doesnt work with absolute urls atm
            break;
          }
          httpConfig.url = $spRest.appendQueryString(httpConfig.url, query);
          httpConfig.transformResponse = $spRest.transformResponse;
          return httpConfig;
        }
      };
    return $spRest;
  }
]);