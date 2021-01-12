require('dotenv-safe').config();

const config = require('config');
const fs = require('fs');
const Path = require('path');
const _ = require('lodash');
const mkdirp = require('mkdirp');
const stringifySafe = require('json-stringify-safe');

const { transports } = require('winston');
const spauth = require('node-sp-auth');
const odatafilter = require('odata-filter-builder').ODataFilterBuilder;

const sharepoint = require('./lib/sharepoint');
const logger = require('./lib/logger');
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

  await authenticate(options)
    .then((authresponse) => {
      options.logger.info(`Authenticated Successfully`, loggingOptions);
      options.sharepoint.authheaders = authresponse.headers;
    })
    .catch((autherr) => {
      options.logger.error(`Error:  ${autherr}`, loggingOptions);
    });

  const basequery = {
    Select: ['Title', 'DESCRIPTION', 'MODALITY', 'LAUNCH'],
    Top: 50,
  };

  await sharepoint
    .getAllItems(
      options,
      'EXAMPLE LIST',
      _.merge({}, basequery, {
        Filter: odatafilter()
          .fn('substringof', 'Title', 'Office', true, true)
          .eq('MODALITY', 'WATCH')
          .lt('MINUTES', 20)
          .toString(),
      })
    )
    .then((response) => {
      options.logger.debug(
        `sharepoint.getAllItems Response: ${stringifySafe(response.data)}`,
        loggingOptions
      );
    })
    .catch((err) => {
      options.logger.error(`sharepoint.getAllItems Error:  ${err}`, loggingOptions);
    });

  await sharepoint
    .getAllItems(options, 'EXAMPLE LIST', _.merge({}, basequery))
    .then((response) => {
      options.logger.debug(
        `sharepoint.getAllItems Response: ${stringifySafe(response.data)}`,
        loggingOptions
      );
    })
    .catch((err) => {
      options.logger.error(`sharepoint.getItems Error:  ${err}`, loggingOptions);
    });

  await sharepoint
    .getAllItems(
      options,
      'EXAMPLE LIST',
      _.merge({}, basequery, {
        Filter: odatafilter().eq('MODALITY', 'READ').toString(),
      })
    )
    .then((response) => {
      options.logger.debug(
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
