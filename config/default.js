const moment = require('moment');
const defer = require('config/defer').deferConfig;

const config = {};
config.startTimestamp = moment().utc().format('YYYYMMDD_HHmmss');

// DEBUG Options
config.debug = {};
// One of the supported default logging levels for winston - see https://github.com/winstonjs/winston#logging-levels
config.debug.loggingLevel = 'info';
config.debug.path = 'results';
config.debug.filename = defer((cfg) => {
  return `${cfg.startTimestamp}_results.log`;
});

// Default for for saving the output
config.output = {};
config.output.path = 'results';
config.output.filename = defer((cfg) => {
  return `${cfg.startTimestamp}_results.json`;
});

config.sharepoint = {};
config.sharepoint.url = null;
config.sharepoint.username = null;
config.sharepoint.password = null;

// Global Web Retry Options for promise retry
// see https://github.com/IndigoUnited/node-promise-retry#readme
config.retry_options = {};
config.retry_options.retries = 3;
config.retry_options.minTimeout = 1000;
config.retry_options.maxTimeout = 2000;

module.exports = config;
