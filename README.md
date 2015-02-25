EWS (Exchange Web Services) Client
=============================
A library to connect to exchange and retrieve calendar and contact information. 
This library uses NTLM for authentication with the exchange server.

## Installation
'''npm install ews'''

## Usage
```JavaScript
var ews = require('ews');

var mailbox = 'email@exchangedomain.com'
var exSession = new ews.MSExchange(ntlmuser, ntlmpass, '', ntlmdomain);

exSession.autoDiscover(mailbox)
.then(function() {

      var calendar = new ews.Calendar(exSession, mailbox);

      calendar.getEntries().then(function(entries) {
	      /* do something with the entries */
		  for (var i = 0; i < entries.length; i++) {
                console.log(entries[i].Subject);
                console.log('From: ' + entries[i].Start + ', To: ' + entries[i].End);
      });

});
```

## Tests
'''npm test'''

## Release History
0.0.1 Initial version
