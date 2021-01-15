const $REST = require('gd-sprest');
const promiseRetry = require('promise-retry');
const Axios = require('axios');
const _ = require('lodash');
const stringifySafe = require('json-stringify-safe');
const { accessSafe } = require('access-safe');

/**
 * Call Sharepoint
 *
 * @param {*} options
 * @param {*} sharepointquery
 * @param {Axios} [axiosInstance=Axios] HTTP request client that provides an Axios like interface
 * @returns
 */
const callSharePointODATA = async (options, sharepointquery, axiosInstance = Axios) => {
  return promiseRetry(async (retry, numberOfRetries) => {
    const loggingOptions = {
      label: 'callSharepoint',
    };

    options.logger.debug(`Sharepoint Query: ${stringifySafe(sharepointquery)}`, loggingOptions);

    const axiosConfig = {
      url: sharepointquery.url,
      headers: { ...sharepointquery.headers, ...options.sharepoint.authheaders },
      method: sharepointquery.method === 'GET' ? 'get' : 'post',
      timeout: 2000,
      data: sharepointquery.data,
    };

    options.logger.debug(`Axios Config: ${stringifySafe(axiosConfig)}`, loggingOptions);

    try {
      const response = await axiosInstance.request(axiosConfig);
      options.logger.debug(`Response Headers: ${stringifySafe(response.headers)}`, loggingOptions);
      return response;
    } catch (err) {
      options.logger.warn(
        `Trying to get report. Got Error after Attempt# ${numberOfRetries} : ${err}`,
        loggingOptions
      );
      if (err.response) {
        options.logger.debug(
          `Response Headers: ${stringifySafe(err.response.headers)}`,
          loggingOptions
        );
        options.logger.debug(`Response Body: ${stringifySafe(err.response.data)}`, loggingOptions);
      } else {
        options.logger.debug('No Response Object available', loggingOptions);
      }
      if (numberOfRetries < options.retry_options.retries + 1) {
        retry(err);
      } else {
        options.logger.error('Failed to call Sharepoint', loggingOptions);
      }
      throw err;
    }
  }, options.retry_options);
};

/**
 * Get lists on site
 *
 * @param {*} options The config object details see config/defaults.js
 * @returns {Promise} Promise object which after 2 seconds returns the configuration value message
 */
const getLists = (options) => {
  const { url } = options.sharepoint;

  const spquery = $REST.Web(url).Lists().getInfo();

  return callSharePointODATA(options, spquery);
};

/**
 * Get items from my example list filtered
 *
 * @param {*} options The config object details see config/defaults.js
 * @param {*} sharepointquery The ODATA Query
 * @returns {Promise} Promise object which after 2 seconds returns the configuration value message
 */
const getItems = (options, queryFilter) => {
  const { url, list } = options.sharepoint;

  const defaultFilter = { Top: 5000 };
  const mandatorySelect = ['ID', 'Title'];

  const filter = _.merge({}, defaultFilter, queryFilter);

  // If Select filter ensure ID in list
  if (accessSafe(() => filter.Select.length, 0) > 0) {
    filter.Select = [...new Set([...mandatorySelect, ...filter.Select])];
  }

  const spquery = $REST.Web(url).Lists(list).Items().query(filter).getInfo();

  return callSharePointODATA(options, spquery);
};

/**
 * Loop thru calling the ODATA Items until all items are delivered.
 *
 * @param {*} options The config object details see config/defaults.js
 * @param {*} sharepointquery The ODATA Query
 * @returns {string} json file path
 */
const getAllItems = async (options, sharepointquery) => {
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
          response = await getItems(opts, query);
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
 * Get items from my example list filtered
 *
 * @param {*} options The config object details see config/defaults.js
 * @returns {Promise} Promise object for the call to SharePoint
 */
const addItem = (options, item) => {
  const { url, list } = options.sharepoint;

  const spquery = $REST.Web(url).Lists(list).Items().add(item).getInfo();

  return callSharePointODATA(options, spquery);
};

/**
 * Get Context Info
 *
 * @param {*} options
 * @return {*}
 */
const getContextInfo = (options) => {
  const { url } = options.sharepoint;

  const spquery = $REST.ContextInfo.getWeb(url).getInfo();

  return callSharePointODATA(options, spquery);
};

module.exports = {
  callSharePointODATA,
  getLists,
  getItems,
  getAllItems,
  addItem,
  getContextInfo,
};
