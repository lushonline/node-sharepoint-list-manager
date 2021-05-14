require('dotenv-safe').config();

const Axios = require('axios');
const config = require('config');
const fs = require('fs');
const Path = require('path');
const _ = require('lodash');
const mkdirp = require('mkdirp');
const stringifySafe = require('json-stringify-safe');
const rateLimit = require('axios-rate-limit');
const rax = require('retry-axios');
const { accessSafe } = require('access-safe');
const $REST = require('gd-sprest');

const { transports } = require('winston');
const spauth = require('node-sp-auth');
const odatafilter = require('odata-filter-builder').ODataFilterBuilder;

const logger = require('./lib/logger');
const timingAdapter = require('./lib/timingAdapter');

const { ListClient } = require('./lib/sharepoint/');

const pjson = require('./package.json');

/**
 * Authenticate to Sharepoint
 *
 * @param {*} options
 * @returns Promise
 */
const authenticate = (options) => {
  const { url, username, password } = options.sharepoint;

  return spauth.getAuth(url, {
    username,
    password,
    online: true,
  });
};

/**
 * Perform the task
 *
 * @param {*} options
 * @returns
 */
const main = async (configOptions) => {
  const loggingOptions = {
    label: 'main',
  };

  const options = configOptions || null;

  options.logger = logger;

  if (_.isNull(options)) {
    options.logger.error('Invalid configuration', loggingOptions);
    return false;
  }

  // Create logging folder if one does not exist
  if (!_.isNull(options.debug.path)) {
    if (!fs.existsSync(options.debug.path)) {
      mkdirp.sync(options.debug.path);
    }
  }

  // Create output folder if one does not exist
  if (!_.isNull(options.output.path)) {
    if (!fs.existsSync(options.output.path)) {
      mkdirp.sync(options.output.path);
    }
  }

  // Add logging to a file
  options.logger.add(
    new transports.File({
      filename: Path.join(options.debug.path, options.debug.filename),
      options: {
        flags: 'w',
      },
    })
  );
  options.logger.info(`Start ${pjson.name} - v${pjson.version}`, loggingOptions);

  options.logger.debug(`Options: ${stringifySafe(options)}`, loggingOptions);

  options.logger.info('Running task', loggingOptions);

  // Create an axios instance that this will allow us to replace
  // with ratelimiting
  // see https://github.com/aishek/axios-rate-limit
  const axiosInstance = rateLimit(Axios.create({ adapter: timingAdapter }), options.ratelimit);

  // Add Axios Retry
  // see https://github.com/JustinBeckwith/retry-axios
  axiosInstance.defaults.raxConfig = _.merge(
    {},
    {
      instance: axiosInstance,
      // You can detect when a retry is happening, and figure out how many
      // retry attempts have been made
      onRetryAttempt: (err) => {
        const raxcfg = rax.getConfig(err);
        logger.warn(
          `CorrelationId: ${err.config.correlationid}. Retry attempt #${raxcfg.currentRetryAttempt}`,
          {
            label: 'onRetryAttempt',
          }
        );
      },
      // Override the decision making process on if you should retry
      shouldRetry: (err) => {
        const cfg = rax.getConfig(err);
        // ensure max retries is always respected
        if (cfg.currentRetryAttempt >= cfg.retry) {
          logger.warn(`CorrelationId: ${err.config.correlationid}. Maximum retries reached.`, {
            label: `shouldRetry`,
          });
          return false;
        }

        // ensure max retries for NO RESPONSE errors is always respected
        if (cfg.currentRetryAttempt >= cfg.noResponseRetries) {
          logger.warn(
            `CorrelationId: ${err.config.correlationid}. Maximum retries reached for No Response Errors.`,
            {
              label: `shouldRetry`,
            }
          );
          return false;
        }

        // Always retry if response was not JSON
        if (err.message.includes('Request did not return JSON')) {
          logger.warn(
            `CorrelationId: ${err.config.correlationid}. Request did not return JSON. Retrying.`,
            {
              label: `shouldRetry`,
            }
          );
          return true;
        }

        // Handle the request based on your other config options, e.g. `statusCodesToRetry`
        if (rax.shouldRetryRequest(err)) {
          return true;
        }

        logger.error(`CorrelationId: ${err.config.correlationid}. None retryable error.`, {
          label: `shouldRetry`,
        });
        return false;
      },
    },
    options.rax
  );
  rax.attach(axiosInstance);

  await authenticate(options)
    .then((authresponse) => {
      options.logger.info(`Authenticated Successfully`, loggingOptions);
      options.sharepoint.authheaders = authresponse.headers;
    })
    .catch((autherr) => {
      options.logger.error(`Error:  ${autherr}`, loggingOptions);
    });

  const listClient = new ListClient(options, axiosInstance);
  await listClient.getContextInfo();

  await listClient
    .getListInfo()
    .then((response) => {
      options.logger.debug(
        `ListClient.getListInfo Response: ${stringifySafe(response.data.d)}`,
        loggingOptions
      );
    })
    .catch((err) => {
      const message = accessSafe(() => JSON.stringify(err.response.data), err.message);
      options.logger.error(`ListClient.getListInfo Error:  ${message}`, loggingOptions);
    });

  await listClient
    .upsertItems(
      [
        { Title: 'Martin22', DESCRIPTION: 'Test22', LANGUAGE: 'Test2' },
        { Title: 'Martin33', DESCRIPTION: 'Test33', LANGUAGE: 'Test3' },
      ],
      ['Title']
    )
    .then((response) => {
      options.logger.debug(
        `ListClient.upsertItems Response: ${stringifySafe(response)}`,
        loggingOptions
      );
    })
    .catch((err) => {
      const message = accessSafe(() => JSON.stringify(err.response.data), err.message);
      options.logger.error(`ListClient.upsertItems Error:  ${message}`, loggingOptions);
    });

  /* await listClient
    .createList()
    .then((response) => {
      options.logger.info(
        `ListClient.createList Response: ${stringifySafe(response)}`,
        loggingOptions
      );
    })
    .catch((err) => {
      options.logger.error(`ListClient.createList Error:  ${err}`, loggingOptions);
    });

  await listClient
    .addField({
      name: 'Martin Test Note1',
      title: 'Martin Test Note1',
      description: 'test',
      type: $REST.Helper.SPCfgFieldType.Note,
      noteType: $REST.SPTypes.FieldNoteType.EnhancedRichText,
    })
    .then((response) => {
      options.logger.info(
        `ListClient.addField Response: ${stringifySafe(response)}`,
        loggingOptions
      );
    })
    .catch((err) => {
      options.logger.error(`ListClient.addField Error:  ${err}`, loggingOptions);
    });
*/
  const basequery = {};

  // Example Filter Query
  // Get records where title begins with
  // { Filter: odatafilter().fn('substringof', 'Title', 'Martin', true, true).toString() };
  // By specific modality
  // {  Filter: odatafilter().eq('MODALITY', 'READ').toString() }

  await listClient
    .getAllItems(
      _.merge({}, basequery, {
        Filter: odatafilter().fn('substringof', 'Title', 'Martin', true, true).toString(),
      })
    )
    .then((response) => {
      options.logger.info(
        `sharepoint.getAllItems Response: ${stringifySafe(response.data)}`,
        loggingOptions
      );
    })
    .catch((err) => {
      options.logger.error(`sharepoint.getAllItems Error:  ${err}`, loggingOptions);
    });

  options.logger.info(`End ${pjson.name} - v${pjson.version}`, loggingOptions);
  return true;
};

try {
  main(config);
} catch (error) {
  throw new Error(`A problem occurred during configuration. ${error.message}`);
}
