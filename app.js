require('dotenv-safe').config();
const $REST = require('gd-sprest');
const config = require('config');
const fs = require('fs');
const Path = require('path');
const _ = require('lodash');
const mkdirp = require('mkdirp');
const stringifySafe = require('json-stringify-safe');
const { accessSafe } = require('access-safe');

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

  await sharepoint
    .getContextInfo(options)
    .then((response) => {
      options.sharepoint.authheaders['X-RequestDigest'] = accessSafe(
        () => response.data.d.GetContextWebInformation.FormDigestValue,
        null
      );
      options.logger.debug(
        `sharepoint.getContextInfo Response: ${stringifySafe(response)}`,
        loggingOptions
      );
    })
    .catch((err) => {
      options.logger.error(`sharepoint.getContextInfo Error:  ${err}`, loggingOptions);
    });

/*   await sharepoint
    .createGenericList(options, 'Martin123456')
    .then((response) => {
      options.logger.info(
        `sharepoint.createGenericList Response: ${stringifySafe(response)}`,
        loggingOptions
      );
    })
    .catch((err) => {
      options.logger.error(`sharepoint.createGenericList Error:  ${err}`, loggingOptions);
    }); */

  await sharepoint
    .addFieldToList(options, 'Martin123456', {
      Title: 'Description',
      FieldTypeKind: $REST.SPTypes.FieldType.Text,
    })
    .then((response) => {
      options.logger.info(
        `sharepoint.addFieldToList Response: ${stringifySafe(response)}`,
        loggingOptions
      );
    })
    .catch((err) => {
      options.logger.error(`sharepoint.addFieldToList Error:  ${err}`, loggingOptions);
    });

  /*   await sharepoint
    .addItem(options, { Title: 'Martin' })
    .then((response) => {
      options.logger.debug(
        `sharepoint.addItem Response: ${stringifySafe(response)}`,
        loggingOptions
      );
    })
    .catch((err) => {
      options.logger.error(`sharepoint.addItem Error:  ${err}`, loggingOptions);
    }); */

  /*   await sharepoint
    .addItems(options, [{ Title: 'Martin1', DESCRIPTION: 'Test' }, { Title: 'Martin2' }])
    .then((response) => {
      const responses = [...response];
      _.forEach(responses, (item) => {
        options.logger.debug(
          `sharepoint.addItems Response: ${stringifySafe(item.data)}`,
          loggingOptions
        );
      });
    })
    .catch((err) => {
      options.logger.error(`sharepoint.addItems Error:  ${err}`, loggingOptions);
    }); */

  const basequery = {};

  await sharepoint
    .upsertItems(
      options,
      [
        { Title: 'Martin2', DESCRIPTION: 'Test2', LANGUAGE: 'Test2' },
        { Title: 'Martin3', DESCRIPTION: 'Test3', LANGUAGE: 'Test3' },
      ],
      ['Title']
    )
    .then((response) => {
      options.logger.debug(
        `sharepoint.upsertItems Response: ${stringifySafe(response.data)}`,
        loggingOptions
      );
    })
    .catch((err) => {
      const message = accessSafe(() => JSON.stringify(err.response.data), err.message);
      options.logger.error(`sharepoint.upsertItems Error:  ${message}`, loggingOptions);
    });

  /*   await sharepoint
    .getAllItems(
      options,
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
    }); */

  await sharepoint
    .getAllItems(options, _.merge({}, basequery))
    .then((response) => {
      options.logger.debug(
        `sharepoint.getAllItems Response: ${stringifySafe(response.data)}`,
        loggingOptions
      );
      /*       const test = _.map(
        accessSafe(() => response.data.d.results, []),
        'UUID'
      ); */
    })
    .catch((err) => {
      options.logger.error(`sharepoint.getItems Error:  ${err}`, loggingOptions);
    });

  await sharepoint
    .getAllItems(
      options,
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
