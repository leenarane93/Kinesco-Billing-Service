const { createLogger, format, transports } = require('winston');
let today = new Date()
let logdate = today.toISOString().split('T')[0];
module.exports = createLogger({
transports:
    new transports.File({
    filename: 'logs/' + logdate +'.log',
    format:format.combine(
        format.timestamp({format: 'MMM-DD-YYYY HH:mm:ss'}),
        format.align(),
        format.printf(info => `${info.level}: ${[info.timestamp]}: ${info.message}`),
    )}),
});