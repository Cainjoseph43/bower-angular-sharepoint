'use strict';
/**
 * @ngdoc overview
 * @name ExpertsInside.SharePoint
 *
 * @description
 *
 * # ExpertsInside.SharePoint
 *
 * The `ExpertsInside.SharePoint` module provides a high level abstraction for
 * the SharePoint 2013 REST API.
 *
 *
 * ## $spList
 *
 * Interaction with SharePoint Lists similiar to $ngResource.
 * See {@link ExpertsInside.SharePoint.$spList `$spList`} for usage.
 */
angular.module('ExpertsInside.SharePoint', ['ng']).run([
  '$window',
  '$log',
  function ($window, $log) {
    if (angular.isUndefined($window.ShareCoffee)) {
      $log.error('angular-sharepoint requires ShareCoffee to do its job. ' + 'Please include ShareCoffe.js in your document');
    }
  }
]);
angular.module('ExpertsInside.SharePoint').factory('$spConvert', function () {
  'use strict';
  var assertType = function (type, obj) {
    if (!angular.isObject(obj.__metadata) || obj.__metadata.type !== type) {
      throw $spConvertMinErr('badargs', 'expected argument to be of type {0}.', type);
    }
  };
  var $spConvertMinErr = angular.$$minErr('$spConvert');
  var $spConvert = {
      spKeyValue: function (keyValue) {
        assertType('SP.KeyValue', keyValue);
        var value = keyValue.Value;
        switch (keyValue.ValueType) {
        case 'Edm.Double':
        case 'Edm.Float':
          return parseFloat(value);
        case 'Edm.Int16':
        case 'Edm.Int32':
        case 'Edm.Int64':
        case 'Edm.Byte':
          return parseInt(value, 10);
        case 'Edm.Boolean':
          return value === 'true';
        default:
          return value;
        }
      },
      spKeyValueArray: function (keyValues) {
        var result = {};
        for (var i = 0, l = keyValues.length; i < l; i += 1) {
          var keyValue = keyValues[i];
          var key = keyValue.Key;
          result[key] = $spConvert.spKeyValue(keyValue);
        }
        return result;
      },
      spSimpleDataRow: function (row) {
        assertType('SP.SimpleDataRow', row);
        var cells = row.Cells.results || [];
        return $spConvert.spKeyValueArray(cells);
      },
      spSimpleDataTable: function (table) {
        assertType('SP.SimpleDataTable', table);
        var result = [];
        var rows = table.Rows.results || [];
        for (var i = 0, l = rows.length; i < l; i += 1) {
          var row = rows[i];
          result.push($spConvert.spSimpleDataRow(row));
        }
        return result;
      },
      searchResult: function (searchResult) {
        assertType('Microsoft.Office.Server.Search.REST.SearchResult', searchResult);
        var primaryQueryResult = searchResult.PrimaryQueryResult;
        var result = {
            elapsedTime: searchResult.ElapsedTime,
            spellingSuggestion: searchResult.SpellingSuggestion,
            properties: $spConvert.spKeyValueArray(searchResult.Properties.results),
            primaryQueryResult: {
              queryId: primaryQueryResult.QueryId,
              queryRuleId: primaryQueryResult.QueryRuleId,
              relevantResults: $spConvert.spSimpleDataTable(primaryQueryResult.RelevantResults.Table),
              customResults: primaryQueryResult.CustomResults !== null ? $spConvert.spSimpleDataTable(primaryQueryResult.CustomResults.Table) : null,
              refinementResults: primaryQueryResult.RefinementResults !== null ? $spConvert.spSimpleDataTable(primaryQueryResult.RefinementResults.Table) : null,
              specialTermResults: primaryQueryResult.SpecialTermResults !== null ? $spConvert.spSimpleDataTable(primaryQueryResult.SpecialTermResults.Table) : null
            }
          };
        return result;
      },
      suggestResult: function (suggestResult) {
        // TODO implement
        return suggestResult;
      },
      userResult: function (userResult) {
        assertType('SP.UserProfiles.PersonProperties', userResult);
        var result = {
            accountName: userResult.AccountName,
            displayName: userResult.DisplayName,
            email: userResult.Email,
            isFollowed: userResult.IsFollowed,
            personalUrl: userResult.PersonalUrl,
            pictureUrl: userResult.PictureUrl,
            profileProperties: $spConvert.spKeyValueArray(userResult.UserProfileProperties),
            title: userResult.Title,
            userUrl: userResult.UserUrl
          };
        return result;
      },
      capitalize: function (str) {
        if (angular.isUndefined(str) || str === null) {
          return null;
        }
        return str.charAt(0).toUpperCase() + str.slice(1);
      }
    };
  return $spConvert;
});
/**
 * @ngdoc service
 * @name ExpertsInside.SharePoint.$spList
 * @requires ExpertsInside.SharePoint.$spRest
 * @requires ExpertsInside.SharePoint.$spConvert
 *
 * @description A factory which creates a list item resource object that lets you interact with
 *   SharePoint List Items via the SharePoint REST API.
 *
 *   The returned list item object has action methods which provide high-level behaviors without
 *   the need to interact with the low level $http service.
 *
 * @param {string} title The title of the SharePoint List (case-sensitive).
 *
 * @param {Object=} listOptions Hash with custom options for this List. The following options are
 *   supported:
 *
 *   - **`readOnlyFields`** - {Array.{string}=} - Array of field names that will be exlcuded
 *   from the request when saving an item back to SharePoint
 *   - **`query`** - {Object=} - Default query parameter used by each action. Can be
 *   overridden per action. See {@link ExpertsInside.SharePoint.$spList query} for details.
 *
 * @return {Object} A list item "class" object with methods for the default set of resource actions.
 *
 * # List Item class
 *
 * All query parameters accept an object with the REST API query string parameters. Prefixing them with $ is optional.
 *   - **`$select`**
 *   - **`$filter`**
 *   - **`$orderby`**
 *   - **`$top`**
 *   - **`$skip`**
 *   - **`$expand`**
 *   - **`$sort`**
 *
 * ## Methods
 *
 *   - **`get`** - {function(id, query)} - Get a single list item by id.
 *   - **`query`** - {function(query, options)} - Query the list for list items and returns the list
 *     of query results.
 *     `options` supports the following properties:
 *       - **`singleResult`** - {boolean} - Returns and empty object instead of an array. Throws an
 *         error when more than one item is returned by the query.
 *   - **`create`** - {function(item, query)} - Creates a new list item. Throws an error when item is
 *     not an instance of the list item class.
 *   - **`update`** - {function(item, options)} - Updates an existing list item. Throws an error when
 *     item is not an instance of the list item class. Supported options are:
 *       - **`query`** - {Object} - Query parameters for the REST call
 *       - **`force`** - {boolean} - If true, the etag (version) of the item is excluded from the
 *         request and the server does not check for concurrent changes to the item but just 
 *         overwrites it. Use with caution.
 *   - **`save`** - {function(item, options)} - Either creates or updates the item based on its state.
 *     `options` are passed down to `update` and and `options.query` are passed down to `create`.
 *   - **`delete`** - {function(item)} - Deletes the list item. Throws an error when item is not an
 *     instance of the list item class.
 *
 * @example
 *
 * # Todo List
 *
 * ## Defining the Todo class
 * ```js
     var Todo = $spList('Todo', {
       query: ['Id', 'Title', 'Completed']
     );
 * ```
 *
 * ## Queries
 *
 * ```js
     // We can retrieve all list items from the server.
     var todos = Todo.query();

    // Or retrieve only the uncompleted todos.
    var todos = Todo.query({
      filter: 'Completed eq 0'
    });

    // Queries that are used in more than one place or those accepting a parameter can be defined 
    // as a function on the class
    Todo.addNamedQuery('uncompleted', function() {
      filter: "Completed eq 0"
    });
    var uncompletedTodos = Todo.queries.uncompleted();
    Todo.addNamedQuery('byTitle', function(title) {
      filter: "Title eq " + title
    });
    var fooTodo = Todo.queries.byTitle('Foo');
 * ```
 */
angular.module('ExpertsInside.SharePoint').factory('$spList', [
  '$spRest',
  '$http',
  '$spConvert',
  function ($spRest, $http, $spConvert) {
    'use strict';
    var $spListMinErr = angular.$$minErr('$spList');
    function listFactory(title, listOptions) {
      if (!angular.isString(title) || title === '') {
        throw $spListMinErr('badargs', 'title must be a nen-empty string.');
      }
      if (!angular.isObject(listOptions)) {
        listOptions = {};
      }
      var normalizedTitle = $spConvert.capitalize(title.replace(/[^A-Za-z0-9 ]/g, '').replace(/\s/g, '_x0020_'));
      var className = $spConvert.capitalize(normalizedTitle.replace(/_x0020/g, '').replace(/^\d+/, ''));
      var listItemType = 'SP.Data.' + normalizedTitle + 'ListItem';
      // Constructor function for List dynamically generated List class
      var List = function () {
          // jshint evil:true
          var script = ' (function() {                     ' + '   function {{List}}(data) {       ' + '     this.__metadata = {           ' + '       type: \'' + listItemType + '\'' + '     };                            ' + '     angular.extend(this, data);   ' + '   }                               ' + '   return {{List}};                ' + ' })();                             ';
          return eval(script.replace(/{{List}}/g, className));
        }();
      List.$title = title;
      /**
       * Web relative list url
       * @private
       */
      List.$$relativeUrl = 'web/lists/getByTitle(\'' + title + '\')';
      /**
       * Is this List in the host web?
       * @private
       */
      List.$$inHostWeb = !!listOptions.inHostWeb;
      /**
       * Decorate the result with $promise and $resolved
       * @private
       */
      List.$$decorateResult = function (result, httpConfig) {
        if (!angular.isArray(result) && !(result instanceof List)) {
          result = new List(result);
        }
        if (angular.isUndefined(result.$resolved)) {
          result.$resolved = false;
        }
        result.$promise = $http(httpConfig).then(function (response) {
          var data = response.data;
          if (angular.isArray(result) && angular.isArray(data)) {
            angular.forEach(data, function (item) {
              result.push(new List(item));
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
      /**
       *
       * @description Get a single list item by id
       *
       * @param {Number} id Id of the list item
       * @param {Object=} query Additional query properties
       *
       * @return {Object} List item instance
       */
      List.get = function (id, query) {
        if (angular.isUndefined(id) || id === null) {
          throw $spListMinErr('badargs', 'id is required.');
        }
        var result = { Id: id };
        var httpConfig = $spRest.buildHttpConfig(List, 'get', {
            id: id,
            query: query
          });
        return List.$$decorateResult(result, httpConfig);
      };
      /**
       *
       * @description Query for the list for items
       *
       * @param {Object=} query Query properties
       * @param {Object=} options Additional query options.
       *   Accepts the following properties:
       *   - **`singleResult`** - {boolean} - Returns and empty object instead of an array. Throws an
       *     error when more than one item is returned by the query.
       *
       * @return {Array<Object>} Array of list items
       */
      List.query = function (query, options) {
        var result = angular.isDefined(options) && options.singleResult ? {} : [];
        var httpConfig = $spRest.buildHttpConfig(List, 'query', { query: angular.extend({}, List.prototype.$$queryDefaults, query) });
        return List.$$decorateResult(result, httpConfig);
      };
      /**
       *
       * @description Save a new list item on the server.
       *
       * @param {Object=} item Query properties
       * @param {Object=} options Additional query properties.
       *
       * @return {Object} The decorated list item
       */
      List.create = function (item, query) {
        if (!(angular.isObject(item) && item instanceof List)) {
          throw $spListMinErr('badargs', 'item must be a List instance.');
        }
        item.__metadata = angular.extend({ type: listItemType }, item.__metadata);
        var httpConfig = $spRest.buildHttpConfig(List, 'create', {
            item: item,
            query: angular.extend({}, item.$$queryDefaults, query)
          });
        return List.$$decorateResult(item, httpConfig);
      };
      /**
       *
       * @description Update an existing list item on the server.
       *
       * @param {Object=} item the list item
       * @param {Object=} options Additional update properties.
       *   Accepts the following properties:
       *   - **`force`** - {boolean} - Overwrite newer versions on the server.
       *
       * @return {Object} The decorated list item
       */
      List.update = function (item, options) {
        if (!(angular.isObject(item) && item instanceof List)) {
          throw $spListMinErr('badargs', 'item must be a List instance.');
        }
        options = angular.extend({}, options, { item: item });
        var httpConfig = $spRest.buildHttpConfig(List, 'update', options);
        return List.$$decorateResult(item, httpConfig);
      };
      /**
       *
       * @description Update or create a list item on the server.
       *
       * @param {Object=} item the list item
       * @param {Object=} options Options passed to create or update.
       *
       * @return {Object} The decorated list item
       */
      List.save = function (item, options) {
        if (angular.isDefined(item.__metadata) && angular.isDefined(item.__metadata.id)) {
          return this.update(item, options);
        } else {
          var query = angular.isObject(options) ? options.query : undefined;
          return this.create(item, query);
        }
      };
      /**
       *
       * @description Delete a list item on the server.
       *
       * @param {Object=} item the list item
       *
       * @return {Object} The decorated list item
       */
      List.delete = function (item) {
        if (!(angular.isObject(item) && item instanceof List)) {
          throw $spListMinErr('badargs', 'item must be a List instance.');
        }
        var httpConfig = $spRest.buildHttpConfig(List, 'delete', { item: item });
        return List.$$decorateResult(item, httpConfig);
      };
      /**
       * Named queries hash
       */
      List.queries = {};
      /**
       *
       * @description Add a named query to the queries hash
       *
       * @param {Object} name name of the query, used as the function name
       * @param {Function} createQuery callback invoked with the arguments passed to
       *   the created named query that creates the final query object
       * @param {Object=} options Additional query options passed to List.query
       *
       * @return {Array} The query result
       */
      List.addNamedQuery = function (name, createQuery, options) {
        List.queries[name] = function () {
          var query = angular.extend({}, List.prototype.$$queryDefaults, createQuery.apply(List, arguments));
          return List.query(query, options);
        };
        return List;
      };
      List.prototype = {
        $$readOnlyFields: angular.extend([
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
        ], listOptions.readOnlyFields),
        $$queryDefaults: angular.extend({}, listOptions.query),
        $save: function (options) {
          return List.save(this, options).$promise;
        },
        $delete: function () {
          return List.delete(this).$promise;
        },
        $isNew: function () {
          return angular.isUndefined(this.__metadata) || angular.isUndefined(this.__metadata.id);
        }
      };
      return List;
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
          if (angular.isDefined(item.$$readOnlyFields)) {
            angular.forEach(item.$$readOnlyFields, function (readOnlyField) {
              delete payload[readOnlyField];
            });
          }
          return angular.toJson(payload);
        },
        buildHttpConfig: function (list, action, options) {
          var baseUrl = list.$$relativeUrl + '/items';
          var httpConfig = { url: baseUrl };
          if (list.$$inHostWeb) {
            httpConfig.hostWebUrl = ShareCoffee.Commons.getHostWebUrl();
          }
          action = angular.isString(action) ? action.toLowerCase() : '';
          options = angular.isDefined(options) ? options : {};
          var query = angular.isDefined(options.query) ? $spRest.normalizeParams(options.query) : {};
          switch (action) {
          case 'get':
            if (angular.isUndefined(options.id)) {
              throw $spRestMinErr('options:get', 'options must have an id');
            }
            httpConfig.url += '(' + options.id + ')';
            httpConfig = ShareCoffee.REST.build.read.for.angularJS(httpConfig);
            break;
          case 'query':
            httpConfig = ShareCoffee.REST.build.read.for.angularJS(httpConfig);
            break;
          case 'create':
            if (angular.isUndefined(options.item)) {
              throw $spRestMinErr('options:create', 'options must have an item');
            }
            if (angular.isDefined(query)) {
              delete query.$expand;
            }
            httpConfig.payload = $spRest.createPayload(options.item);
            httpConfig = ShareCoffee.REST.build.create.for.angularJS(httpConfig);
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
            httpConfig.url += '(' + options.item.Id + ')';
            httpConfig.payload = $spRest.createPayload(options.item);
            httpConfig.eTag = !options.force && angular.isDefined(options.item.__metadata) ? options.item.__metadata.etag : null;
            httpConfig = ShareCoffee.REST.build.update.for.angularJS(httpConfig);
            break;
          case 'delete':
            if (angular.isUndefined(options.item)) {
              throw $spRestMinErr('options:delete', 'options must have an item');
            }
            if (angular.isUndefined(options.item.__metadata)) {
              throw $spRestMinErr('options:delete', 'options.item must have __metadata');
            }
            httpConfig.url += '(' + options.item.Id + ')';
            httpConfig = ShareCoffee.REST.build.delete.for.angularJS(httpConfig);
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
/**
 * @ngdoc service
 * @name ExpertsInside.SharePoint.$spSearch
 * @requires ExpertsInside.SharePoint.$spRest
 * @requires ExpertsInside.SharePoint.$spConvert
 *
 * @description Query the Search via REST API
 *
 */
angular.module('ExpertsInside.SharePoint').factory('$spSearch', [
  '$http',
  '$spRest',
  '$spConvert',
  function ($http, $spRest, $spConvert) {
    'use strict';
    var $spSearchMinErr = angular.$$minErr('$spSearch');
    var search = {
        $createQueryProperties: function (searchType, properties) {
          var queryProperties;
          switch (searchType) {
          case 'postquery':
            queryProperties = new ShareCoffee.PostQueryProperties();
            break;
          case 'suggest':
            queryProperties = new ShareCoffee.SuggestProperties();
            break;
          default:
            queryProperties = new ShareCoffee.QueryProperties();
            break;
          }
          return angular.extend(queryProperties, properties);
        },
        $decorateResult: function (result, httpConfig) {
          if (angular.isUndefined(result.$resolved)) {
            result.$resolved = false;
          }
          result.$raw = null;
          result.$promise = $http(httpConfig).then(function (response) {
            var data = response.data;
            if (angular.isObject(data)) {
              if (angular.isDefined(data.query)) {
                result.$raw = data.query;
                angular.extend(result, $spConvert.searchResult(data.query));
              } else if (angular.isDefined(data.suggest)) {
                result.$raw = data.suggest;
                angular.extend(result, $spConvert.suggestResult(data.suggest));
              }
            }
            if (angular.isUndefined(result.$raw)) {
              throw $spSearchMinErr('badresponse', 'Response does not contain a valid search result.');
            }
            result.$resolved = true;
            return result;
          });
          return result;
        },
        query: function (properties) {
          properties = angular.extend({}, properties);
          var searchType = properties.searchType;
          delete properties.searchType;
          var queryProperties = search.$createQueryProperties(searchType, properties);
          var httpConfig = ShareCoffee.REST.build.read.for.angularJS(queryProperties);
          httpConfig.transformResponse = $spRest.transformResponse;
          var result = {};
          return search.$decorateResult(result, httpConfig);
        },
        postquery: function (properties) {
          properties = angular.extend(properties, { searchType: 'postquery' });
          return search.query(properties);
        },
        suggest: function (properties) {
          properties = angular.extend(properties, { searchType: 'suggest' });
          return search.query(properties);
        }
      };
    return search;
  }
]);
/**
 * @ngdoc service
 * @name ExpertsInside.SharePoint.$spUser
 * @requires ExpertsInside.SharePoint.$spRest
 * @requires ExpertsInside.SharePoint.$spConvert
 *
 * @description Load user information via UserProfiles REST API
 *
 */
angular.module('ExpertsInside.SharePoint').factory('$spUser', [
  '$http',
  '$spRest',
  '$spConvert',
  function ($http, $spRest, $spConvert) {
    'use strict';
    var $spUserMinErr = angular.$$minErr('$spUser');
    var $spUser = {
        $decorateResult: function (result, httpConfig) {
          if (angular.isUndefined(result.$resolved)) {
            result.$resolved = false;
          }
          result.$raw = null;
          result.$promise = $http(httpConfig).then(function (response) {
            var data = response.data;
            if (angular.isDefined(data)) {
              result.$raw = data;
              angular.extend(result, $spConvert.userResult(data));
            } else {
              throw $spUserMinErr('badresponse', 'Response does not contain a valid user result.');
            }
            result.$resolved = true;
            return result;
          });
          return result;
        },
        current: function () {
          var properties = new ShareCoffee.UserProfileProperties(ShareCoffee.Url.GetMyProperties);
          var httpConfig = ShareCoffee.REST.build.read.for.angularJS(properties);
          httpConfig.transformResponse = $spRest.transformResponse;
          var result = {};
          return $spUser.$decorateResult(result, httpConfig);
        },
        get: function (accountName) {
          var properties = new ShareCoffee.UserProfileProperties(ShareCoffee.Url.GetProperties, accountName);
          var httpConfig = ShareCoffee.REST.build.read.for.angularJS(properties);
          httpConfig.transformResponse = $spRest.transformResponse;
          var result = {};
          return $spUser.$decorateResult(result, httpConfig);
        }
      };
    return $spUser;
  }
]);