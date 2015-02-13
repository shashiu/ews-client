EWS (Exchange Web Services) Client
=============================
A library to connect to exchange and retrieve calendar and contact information. 
This library uses NTLM for authentication with the exchange server.

## Installation
npm install ews

## Usage
var ews = require('ews');

var mailbox = 'youremail@exchangedomain.com'
var msExchange = new ews.MSExchange();

msExchange.setAuth(ntlmuser, ntlmpass, '', ntlmdomain);
msExchange.autoDiscover(mailbox)
.then(function() {

      var calendar = new ews.Calendar(session, mailbox);
      calendar.getEntries().then(function(entries) {
	      /* do something with the entries */
      });

});

## Tests
npm test

## Release History
0.0.1 Initial version
