const $REST = require('gd-sprest');
const Axios = require('axios');
const _ = require('lodash');
const { accessSafe } = require('access-safe');
const odatafilter = require('odata-filter-builder').ODataFilterBuilder;
const { v4: uuidv4 } = require('uuid');

const nullLogger = require('../nulllogger');

class BaseClient {
  /**
   * Creates an instance of SharePointClient.
   * @param {*} options
   * @param {*} [axiosInstance=Axios] An Instance of
   * @memberof BaseClient
   */
  constructor(options, axiosInstance = Axios) {
    if (_.isNil(options)) {
      throw new Error('options not specified');
    }

    if (_.isNil(accessSafe(() => options.sharepoint, null))) {
      throw new Error('options.sharepoint not specified');
    }

    if (_.isNil(accessSafe(() => options.sharepoint.url, null))) {
      throw new Error('options.sharepoint.url not specified');
    }

    if (_.isNil(accessSafe(() => options.sharepoint.authheaders, null))) {
      throw new Error('options.sharepoint.authheaders not specified');
    }

    this.sharepoint = options.sharepoint;
    this.logger = options.logger || nullLogger;

    this.axiosInstance = axiosInstance;
  }

  /**
   * Builds a simple ODATA filter where all specified item lookups eq
   *
   * @param {object} item - the object with populated with the data to match
   * @param {string[]} lookups - the item field names
   * @return {string} - the Filter value
   * @memberof SharePointClient
   */
  // eslint-disable-next-line class-methods-use-this
  _buildFilter(item, lookups) {
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
  }

  /**
   * Checks that the request digest header is present
   * if the header isnt present then only READ operations are supported
   *
   * @return {Boolean}
   * @memberof BaseClient
   */
  _requireDigest() {
    const { authheaders } = this.sharepoint;
    if (_.isNil(authheaders['X-RequestDigest'])) {
      throw new Error(
        'X-RequestDigest authentication header not set so READ operations only. Call getContextInfo().'
      );
    }
  }

  /**
   * Is this an axios error.
   *
   * @param {Object} err The error to test.
   * @returns {boolean} True if axios error.
   */
  // eslint-disable-next-line class-methods-use-this
  isAxiosError(err) {
    return err instanceof Error && err.config != null && err.request != null;
  }

  /**
   * This will clean the Axios Error
   *
   * @param {Object} err The error to clean.
   * @return {Object}
   * @memberof BaseClient
   */
  getCleanedAxiosError(err) {
    if (!this.isAxiosError(err)) return err;

    const cleanedAxiosError = new Error(err.message);
    if (err.stack != null) cleanedAxiosError.stack = err.stack;
    if (err.errno != null) cleanedAxiosError.errno = err.errno;
    if (err.code != null) cleanedAxiosError.code = err.code;

    const { data = null, status = null, statusText = null } = err.response;
    cleanedAxiosError.response = { data, status, statusText };

    return cleanedAxiosError;
  }

  /**
   * This will clean the Axios Response
   *
   * @param {Object} response The response to clean.
   * @return {Object}
   * @memberof BaseClient
   */
  // eslint-disable-next-line class-methods-use-this
  getCleanedAxiosResponse(response) {
    if (!accessSafe(() => this.sharepoint.clientdebug, false)) {
      const { data = null, status = null, statusText = null, timings = null } = response;
      return { data, status, statusText, timings };
    }
    return response;
  }

  /**
   * Call Sharepoint ODATA API
   *
   * @param {Object} sprequest
   * @returns {Promise}
   * @memberof SharePointClient
   */
  callSharePointODATA(sprequest) {
    return new Promise((resolve, reject) => {
      const { authheaders, timeout, correlationid } = this.sharepoint;

      const axiosConfig = {
        url: sprequest.url,
        headers: { ...sprequest.headers, ...authheaders },
        method: sprequest.method === 'GET' ? 'get' : 'post',
        timeout: timeout || 2000,
        data: sprequest.data,
        correlationid: correlationid || uuidv4(),
      };

      this.axiosInstance
        .request(axiosConfig)
        .then((response) => {
          resolve(this.getCleanedAxiosResponse(response));
        })
        .catch((err) => {
          reject(this.getCleanedAxiosError(err));
        });
    });
  }

  /**
   * Get Context Info, and set X-RequestDigest Header
   *
   * @returns {Promise} Promise object with ODATA response object
   * @memberof SharePointClient
   */
  getContextInfo() {
    return new Promise((resolve, reject) => {
      const { url } = this.sharepoint;

      const sprequest = $REST.ContextInfo.getWeb(url).getInfo();

      this.callSharePointODATA(sprequest)
        .then((response) => {
          // Extract the X-RequestDigest to enable WRITE operations on list
          this.sharepoint.authheaders['X-RequestDigest'] = accessSafe(
            () => response.data.d.GetContextWebInformation.FormDigestValue,
            null
          );
          resolve(response);
        })
        .catch((err) => {
          reject(err);
        });
    });
  }

  /**
   * Alias for getContextInfo()
   *
   * @returns {Promise} Promise object with ODATA response object
   * @memberof BaseClient
   */
  init() {
    return this.getContextInfo();
  }
}

module.exports = {
  BaseClient,
};
