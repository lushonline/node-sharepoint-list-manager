const $REST = require('gd-sprest');
const Axios = require('axios');
const _ = require('lodash');
const stringifySafe = require('json-stringify-safe');
const { accessSafe } = require('access-safe');
const odatafilter = require('odata-filter-builder').ODataFilterBuilder;
const { v4: uuidv4 } = require('uuid');

/**
 * Call Sharepoint ODATA API
 *
 * @param {*} options
 * @param {*} sharepointquery
 * @param {Axios} [axiosInstance=Axios] HTTP request client that provides an Axios like interface
 * @returns {Promise}
 */
const callSharePointODATA = (options, sharepointquery, axiosInstance = Axios) => {
  return new Promise((resolve, reject) => {
    const opts = _.cloneDeep(options);

    const axiosConfig = {
      url: sharepointquery.url,
      headers: { ...sharepointquery.headers, ...opts.sharepoint.authheaders },
      method: sharepointquery.method === 'GET' ? 'get' : 'post',
      timeout: opts.sharepoint.timeout || 2000,
      data: sharepointquery.data,
      correlationid: opts.sharepoint.correlationid || uuidv4(),
    };

    axiosInstance
      .request(axiosConfig)
      .then((response) => {
        resolve(response);
      })
      .catch((err) => {
        reject(err);
      });
  });
};

/**
 * Create Generic list
 *
 * @param {object} options The config object details see config/defaults.js
 * @param {string} title The list title
 * @param {Axios} [axiosInstance=null] HTTP request client that provides an Axios like interface
 * @returns {Promise} Promise object with Axios response object
 */
const createGenericList = (options, title, axiosInstance = null) => {
  const { url } = options.sharepoint;

  if (title == null) {
    throw new Error('No title specified');
  }

  const sprequest = $REST
    .Web(url)
    .Lists()
    .add({
      BaseTemplate: $REST.SPTypes.ListTemplateType.GenericList,
      Title: title,
    })
    .getInfo();

  return callSharePointODATA(options, sprequest, axiosInstance);
};

/**
 * Create Generic list
 *
 * @param {object} options The config object details see config/defaults.js
 * @param {string} listtitle The list title
 * @param {object} field The field definition
 * @param {Axios} [axiosInstance=null] HTTP request client that provides an Axios like interface
 * @returns {Promise} Promise object with Axios response object
 */
const addFieldToList = (options, title, field, axiosInstance = null) => {
  const { url } = options.sharepoint;

  if (title == null) {
    throw new Error('No title specified');
  }

  const sprequest = $REST.Web(url).Lists(title).Fields().addField(field).getInfo();

  return callSharePointODATA(options, sprequest, axiosInstance);
};

/**
 * Get lists on site
 *
 * @param {*} options The config object details see config/defaults.js
 * @param {Axios} [axiosInstance=null] HTTP request client that provides an Axios like interface
 * @returns {Promise} Promise object with Axios response object
 */
const getLists = (options, axiosInstance = null) => {
  const { url } = options.sharepoint;

  const sprequest = $REST.Web(url).Lists().getInfo();

  return callSharePointODATA(options, sprequest, axiosInstance);
};

/**
 * Get list by title
 *
 * @param {object} options The config object details see config/defaults.js
 * @param {string} title The list title
 * @param {Axios} [axiosInstance=null] HTTP request client that provides an Axios like interface
 * @returns {Promise} Promise object with Axios response object
 */
const getListByTitle = (options, title, axiosInstance = null) => {
  const { url } = options.sharepoint;

  if (title == null) {
    throw new Error('No title specified');
  }

  const sprequest = $REST.Web(url).Lists(title).getInfo();

  return callSharePointODATA(options, sprequest, axiosInstance);
};

/**
 * Get items from list
 *
 * @param {*} options The config object details see config/defaults.js
 * @param {*} sharepointquery The ODATA Query
 * @param {Axios} [axiosInstance=null] HTTP request client that provides an Axios like interface
 * @returns {Promise} Promise object with Axios response object
 */
const getItems = (options, sharepointquery, axiosInstance = null) => {
  const { url, list } = options.sharepoint;

  const defaultFilter = { Top: 5000 };
  const filter = _.merge({}, defaultFilter, sharepointquery);

  // ensure ID always in select filter
  if (accessSafe(() => filter.Select.length, 0) > 0) {
    filter.Select = [...new Set([filter.Select])];
  }

  const sprequest = $REST.Web(url).Lists(list).Items().query(filter).getInfo();

  return callSharePointODATA(options, sprequest, axiosInstance);
};

/**
 * Loop thru calling the ODATA Items until all items are delivered.
 *
 * @param {*} options The config object details see config/defaults.js
 * @param {*} sharepointquery The ODATA Query
 * @param {Axios} [axiosInstance=null] HTTP request client that provides an Axios like interface
 * @returns {Promise} Promise object with Axios response object
 */
const getAllItems = async (options, sharepointquery, axiosInstance = null) => {
  // eslint-disable-next-line no-async-promise-executor
  return new Promise(async (resolve, reject) => {
    const loggingOptions = {
      label: 'getAllItems',
    };

    const opts = options;
    opts.logcount = opts.logcount || 500;

    const query = _.omit(sharepointquery, ['custom']);
    let keepGoing = true;
    let downloadedRecords = 0;
    let allrecords = [];
    let lastId = null;

    try {
      while (keepGoing) {
        let response = null;
        opts.logger.info(`Sharepoint Query: ${stringifySafe(query)}`, loggingOptions);

        try {
          // eslint-disable-next-line no-await-in-loop
          response = await getItems(opts, query, axiosInstance);
        } catch (err) {
          opts.logger.error('ERROR: trying to download results', loggingOptions);
          keepGoing = false;
          reject(err);
          break;
        }

        const countRecords = accessSafe(() => response.data.d.results.length, 0);
        const records = accessSafe(() => response.data.d.results, []);
        // eslint-disable-next-line no-underscore-dangle
        const morerecords = accessSafe(() => !_.isEmpty(response.data.d.__next), false);

        if (countRecords > 0) {
          downloadedRecords += countRecords;

          opts.logger.info(
            `Items Downloaded ${downloadedRecords.toLocaleString()}`,
            loggingOptions
          );
          allrecords = allrecords.concat(...records);
        }

        if (morerecords) {
          // Get last ID
          // $skiptoken=Paged=TRUE&p_ID=11
          lastId = allrecords.pop().ID;
          query.Custom = `$skiptoken=Paged%3dTRUE%26p_ID%3d${lastId}`;
          keepGoing = true;
        } else {
          keepGoing = false;
          resolve({ data: { d: { results: allrecords } } });
        }
      }
    } catch (error) {
      reject(error);
    }
  });
};

/**
 * Add item to list
 *
 * @param {object} options The config object details see config/defaults.js
 * @param {object} item The item to add
 * @param {Axios} [axiosInstance=null] HTTP request client that provides an Axios like interface
 * @returns {Promise} Promise object with Axios response object
 */
const addItem = (options, item, axiosInstance = null) => {
  const { url, list } = options.sharepoint;

  const sprequest = $REST.Web(url).Lists(list).Items().add(item).getInfo();

  return callSharePointODATA(options, sprequest, axiosInstance);
};

/**
 * Update item to list
 *
 * @param {*} options The config object details see config/defaults.js
 * @param {object} item The item to add, it must have ID
 * @param {Axios} [axiosInstance=null] HTTP request client that provides an Axios like interface
 * @returns {Promise} Promise object with Axios response object
 */
const updateItem = (options, item, axiosInstance = null) => {
  const { url, list } = options.sharepoint;

  if (item.ID == null) {
    throw new Error('No ID specified');
  }

  const sprequest = $REST
    .Web(url)
    .Lists(list)
    .Items(item.ID)
    .update(_.omit(item, ['ID']))
    .getInfo();

  return callSharePointODATA(options, sprequest, axiosInstance);
};

/**
 * Upsert item to list
 *
 * @param {*} options The config object details see config/defaults.js
 * @param {object} item The item to add
 * @param {*} sharepointquery The ODATA Query to check if item already exists
 * @param {Axios} [axiosInstance=null] HTTP request client that provides an Axios like interface
 * @returns {Promise} Promise object with Axios response object
 */
const upsertItem = async (options, item, sharepointquery, axiosInstance = null) => {
  const defaultFilter = { Top: 1, Filter: odatafilter().eq('ID', item.ID).toString() };
  const filter = _.merge({}, defaultFilter, sharepointquery);

  // Check for item
  await getItems(options, filter, axiosInstance).then((response) => {
    if (accessSafe(() => response.data.d.results, []).length === 0) {
      // Update
      return addItem(options, _.omit(item, ['ID']), axiosInstance);
    }
    const updateitem = _.merge({}, { ID: response.data.d.results[0].ID }, item);
    return updateItem(options, updateitem, axiosInstance);
  });
};

/**
 * Builds asimple ODATA filter where all specified item lookups eq
 *
 * @param {object} item - the object with populated with the data to match
 * @param {string[]} lookups - the item field names
 * @return {string} - the Filter value
 */
const buildFilter = (item, lookups) => {
  const filter = odatafilter();

  if (lookups.length === 0) {
    return '';
  }

  lookups.forEach((lookupitem) => {
    if (!_.isNull(lookupitem)) {
      filter.eq(lookupitem, _.get(item, lookupitem));
    }
  });

  return filter.toString();
};

/**
 * Upsert all item to list, matches on lookupfield
 *
 * @param {*} options The config object details see config/defaults.js
 * @param {object[]} item The items to add
 * @param {string|string[]} lookupitems The field[s] on teh item we try and match
 * @param {Axios} [axiosInstance=null] HTTP request client that provides an Axios like interface
 * @returns {Promise} Promise object with array of Axios response object
 */
const upsertItems = async (options, items, lookupitems, axiosInstance = null) => {
  const promises = [];
  const lookupArray = _.isNil(lookupitems) ? [] : _.castArray(lookupitems);

  _.forEach(items, (item) => {
    promises.push(
      upsertItem(options, item, { Filter: buildFilter(item, lookupArray) }, axiosInstance)
    );
  });

  return Promise.all(promises);
};

/**
 * Add all items to list
 *
 * @param {*} options The config object details see config/defaults.js
 * @param {object[]} items The items to add
 * @param {Axios} [axiosInstance=null] HTTP request client that provides an Axios like interface
 * @returns {Promise} Promise object with array of Axios response object
 */
const addItems = (options, items, axiosInstance = null) => {
  const promises = [];

  _.forEach(items, (item) => {
    promises.push(addItem(options, item, axiosInstance));
  });

  return Promise.all(promises);
};

/**
 * Get Context Info
 *
 * @param {*} options The config object details see config/defaults.js
 * @param {Axios} [axiosInstance=null] HTTP request client that provides an Axios like interface
 * @returns {Promise} Promise object with Axios response object
 */
const getContextInfo = (options, axiosInstance = null) => {
  const { url } = options.sharepoint;

  const sprequest = $REST.ContextInfo.getWeb(url).getInfo();

  return callSharePointODATA(options, sprequest, axiosInstance);
};

module.exports = {
  callSharePointODATA,
  createGenericList,
  addFieldToList,
  getLists,
  getListByTitle,
  getItems,
  getAllItems,
  addItem,
  addItems,
  upsertItem,
  upsertItems,
  updateItem,
  getContextInfo,
};
