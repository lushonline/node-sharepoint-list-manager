const $REST = require('gd-sprest');
const Axios = require('axios');
const _ = require('lodash');
const { accessSafe } = require('access-safe');
const odatafilter = require('odata-filter-builder').ODataFilterBuilder;

const { BaseClient } = require('./BaseClient');

class ListClient extends BaseClient {
  constructor(options, axiosInstance = Axios) {
    super(options, axiosInstance);

    if (_.isNil(accessSafe(() => options.sharepoint.list, null))) {
      throw new Error('options.sharepoint.list not specified');
    }
  }

  /**
   * Create a list
   *
   * @param {$REST.SPTypes.ListTemplateType} [listType=$REST.SPTypes.ListTemplateType.GenericList]
   * @return {Promise}
   * @memberof ListClient
   */
  createList(listType = $REST.SPTypes.ListTemplateType.GenericList) {
    // Write operation so confirm X-RequestDigest configured
    this._requireDigest();

    const { url, list } = this.sharepoint;

    const sprequest = $REST
      .Web(url)
      .Lists()
      .add({
        BaseTemplate: listType,
        Title: list,
      })
      .getInfo();

    return this.callSharePointODATA(sprequest);
  }

  /**
   * Get the list fields
   *
   * @param {*} [filter={}] The ODATA Query
   * @returns {Promise} Promise object with Axios response object
   * @memberof ListClient
   */
  getFields(filter = {}) {
    const { url, list } = this.sharepoint;

    const sprequest = $REST.Web(url).Lists(list).Fields().query(filter).getInfo();

    return this.callSharePointODATA(sprequest);
  }

  /**
   * Add field to list
   *
   * @param {object} fieldInfo The field definition
   * @returns {Promise} Promise object with Axios response object
   */
  addField(fieldInfo) {
    // eslint-disable-next-line no-async-promise-executor
    return new Promise(async (resolve, reject) => {
      // Write operation so confirm X-RequestDigest configured
      this._requireDigest();

      const { url, list } = this.sharepoint;

      const filter = { Filter: odatafilter().eq('Title', fieldInfo.title).toString() };

      this.getFields(filter).then((fieldsResponse) => {
        if (accessSafe(() => fieldsResponse.data.d.results.length, 0) === 0) {
          $REST.Helper.FieldSchemaXML(fieldInfo)
            .then((schemaResponse) => {
              const sprequest = $REST
                .Web(url)
                .Lists(list)
                .Fields()
                .createFieldAsXml(schemaResponse)
                .getInfo();

              this.callSharePointODATA(sprequest)
                .then((response) => resolve(response))
                .catch((err) => reject(err));
            })
            .catch((schemaErr) => reject(schemaErr));
        } else {
          resolve(fieldsResponse);
        }
      });
    });
  }

  /**
   * Get information about the list
   *
   * @param {boolean} [includeFields=false] Include information on the Fields in list
   * @return {Promise}
   * @memberof ListClient
   */
  getListInfo(includeFields = false) {
    const { url, list } = this.sharepoint;

    const filter = includeFields
      ? {
          Expand: ['Fields'],
        }
      : {};

    const sprequest = $REST.Web(url).Lists(list).query(filter).getInfo();

    return this.callSharePointODATA(sprequest);
  }

  /**
   * Get items from list
   *
   * @param {*} [itemFilter={}] The ODATA Query
   * @returns {Promise} Promise object with Axios response object
   * @memberof ListClient
   */
  getItems(itemFilter = {}) {
    const { url, list } = this.sharepoint;

    const defaultFilter = { Top: 1000 };
    const filter = _.merge({}, defaultFilter, itemFilter);

    // ensure ID always in select filter
    if (accessSafe(() => filter.Select.length, 0) > 0) {
      filter.Select = [...new Set([filter.Select])];
    }

    const sprequest = $REST.Web(url).Lists(list).Items().query(filter).getInfo();

    return this.callSharePointODATA(sprequest);
  }

  /**
   * Get item from list by Title
   *
   * @param {string} [title] The Item Title
   * @returns {Promise} Promise object with Axios response object
   * @memberof ListClient
   */
  getItemByTitle(title) {
    if (_.isNil(title)) {
      throw new Error('title not specified');
    }

    const { url, list } = this.sharepoint;

    const filter = { Top: 1, Filter: odatafilter().eq('Title', title).toString() };

    const sprequest = $REST.Web(url).Lists(list).Items().query(filter).getInfo();

    return this.callSharePointODATA(sprequest);
  }

  /**
   * Loop thru calling the ODATA Items until all items are delivered.
   *
   * @param {object} [itemFilter={}] The ODATA Query
   * @returns {Promise} Promise object with Axios response object
   * @memberof ListClient
   */
  async getAllItems(itemFilter = {}) {
    // eslint-disable-next-line no-async-promise-executor
    return new Promise(async (resolve, reject) => {
      const filter = _.omit(itemFilter, ['custom']);
      let keepGoing = true;
      let allrecords = [];
      let lastId = null;

      try {
        while (keepGoing) {
          let response = null;

          try {
            // eslint-disable-next-line no-await-in-loop
            response = await this.getItems(filter);
          } catch (err) {
            keepGoing = false;
            reject(err);
            break;
          }

          const countRecords = accessSafe(() => response.data.d.results.length, 0);
          const records = accessSafe(() => response.data.d.results, []);
          // eslint-disable-next-line no-underscore-dangle
          const morerecords = accessSafe(() => !_.isEmpty(response.data.d.__next), false);

          if (countRecords > 0) {
            allrecords = allrecords.concat(...records);
          }

          if (morerecords) {
            // Get last ID
            // $skiptoken=Paged=TRUE&p_ID=11
            lastId = allrecords.pop().ID;
            filter.Custom = `$skiptoken=Paged%3dTRUE%26p_ID%3d${lastId}`;
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
  }

  /**
   * Add item to list
   *
   * @param {object} item The item to add
   * @returns {Promise} Promise object with Axios response object
   */
  addItem(item) {
    // Write operation so confirm X-RequestDigest configured
    this._requireDigest();

    if (_.isNil(item)) {
      throw new Error('item not specified');
    }

    const { url, list } = this.sharepoint;
    const sprequest = $REST.Web(url).Lists(list).Items().add(item).getInfo();

    return this.callSharePointODATA(sprequest);
  }

  /**
   * Update item to list
   *
   * @param {object} item The item to update, it must have ID
   * @returns {Promise} Promise object with Axios response object
   */
  updateItem(item) {
    // Write operation so confirm X-RequestDigest configured
    this._requireDigest();

    if (_.isNil(item)) {
      throw new Error('item not specified');
    }

    if (accessSafe(() => item.ID, null) == null) {
      throw new Error('No ID specified');
    }

    const { url, list } = this.sharepoint;
    const itemid = item.ID;
    const updateitem = _.omit(item, ['ID']);

    const sprequest = $REST.Web(url).Lists(list).Items(itemid).update(updateitem).getInfo();

    return this.callSharePointODATA(sprequest);
  }

  /**
   * Delete item from list
   *
   * @param {number} itemid The ID number of the item to delete
   * @returns {Promise} Promise object with Axios response object
   */
  deleteItemById(itemid) {
    // Write operation so confirm X-RequestDigest configured
    this._requireDigest();

    if (_.isNil(itemid)) {
      throw new Error('No ID specified');
    }

    const { url, list } = this.sharepoint;
    const sprequest = $REST.Web(url).Lists(list).Items(itemid).delete().getInfo();

    return this.callSharePointODATA(sprequest);
  }

  /**
   * Upsert item to list
   *
   * @param {object} item The item to add/update
   * @param {string|string[]} lookup The field[s] on the item we try and match
   * @returns {Promise} Promise object with Axios response object
   */
  async upsertItem(item, lookup = 'ID') {
    // Write operation so confirm X-RequestDigest configured
    this._requireDigest();

    // eslint-disable-next-line no-async-promise-executor
    return new Promise(async (resolve, reject) => {
      const lookupArray = _.isNil(lookup) ? [] : _.castArray(lookup);

      const defaultFilter = { Top: 1 };
      const queryFilter = { Filter: this._buildFilter(item, lookupArray) };
      const filter = _.merge({}, defaultFilter, queryFilter);

      // Check for item
      await this.getItems(filter)
        .then(async (response) => {
          if (accessSafe(() => response.data.d.results, []).length === 0) {
            // Update
            await this.addItem(_.omit(item, ['ID']))
              .then((addResponse) => {
                resolve(addResponse);
              })
              .catch((addErr) => reject(addErr));
          }
          const updateitem = _.merge({}, { ID: response.data.d.results[0].ID }, item);
          await this.updateItem(updateitem)
            .then((updateResponse) => {
              resolve(updateResponse);
            })
            .catch((updateErr) => reject(updateErr));
        })
        .catch((err) => reject(err));
    });
  }
}

module.exports = {
  ListClient,
};
