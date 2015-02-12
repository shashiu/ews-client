EWS
====
A library to connect to exchange and retrieve calendar and contact information. 
This library uses NTLM for authentication with the exchange server.

## Installation
npm install ews

## Usage
var ews = require('./../ews');
var mailbox = 'youremail@exchangedomain.com'
var msExchange = new ews.MSExchange();
var calendar = new ews.Calendar(msExchange, mailbox);
msExchange.setAuth(ntlmuser, ntlmpass, '', ntlmdomain);
msExchange.autoDiscover(mailbox).then(function()
      calendar.getEntries(roomEmail).then(function(entries) {
      }
}

## Tests
npm test

## Release History
0.0.1 Initial version
