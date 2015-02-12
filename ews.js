/* 
  Exchange web services - a library to connect to MS exchange and get to the calendar and contacts
  Uses NTLM authentication.
*/
exports.MSExchange = require('./lib/msexchange').MSExchange;
exports.Calendar = require('./lib/calendar').Calendar;
exports.Contacts = require('./lib/contact').Contacts;