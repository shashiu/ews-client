var ews = require('./../ews');

// get web credetials from process environment
var user = process.env.EWSUSER;
var pass = process.env.EWSPASS;
var domain = process.env.EWSDOMAIN;
var mailbox = process.env.EWSTESTEMAIL;
if (user == null || pass == null || domain == null || mailbox == null) {
    console.log("System Environment not setup - EWSUSER, EWSPASS, EWSDOMAIN, EWSTESTEMAIL (without quotes)");
    process.exit(1);
}

// setup connection to Exchange
var msExchange = new ews.MSExchange();
var calendar = new ews.Calendar(msExchange);
var contacts = new ews.Contacts(msExchange);

// Autodiscover URL to get to exchange web service
msExchange.setAuth(user, pass, domain);
console.log("Autodiscover");
msExchange.autoDiscover(mailbox)
    .then(function () {
        console.log("Autodiscover Success!");
        // enumerate all room resources
        return contacts.getRoomLists();
    }).then(function (roomLists) {
        if (roomLists == null) {
            console.log('    Room List is null (Exchange Admin hasn\'t configured them)');
        }
        // enumerate calendar for a room
        console.log('Getting calendar entries for next 24 hours for ' + mailbox);
        console.log('----------------------------------------------------------------------------');
        return calendar.getEntries(mailbox);
    }).then(function (entries) {
        if (entries != null) {
            for (var i = 0; i < entries.length; i++) {
                // display the calendar entries for the room
                console.log(entries[i].Subject);
                console.log('    From: ' + entries[i].Start + ', To: ' + entries[i].End)
                if (entries[i].MeetingUrl != null) {
                    // including meeting URL and audio information, if there is some
                    console.log('    Meeting URL:' + entries[i].MeetingUrl);
                    console.log('    AudioOptions:' + entries[i].GTMAudioOptions);
                    console.log('    Attendees:' + entries[i].Attendees);
                }
            }
        }
        console.log('----------------------------------------------------------------------------');
        return contacts.getDetails(mailbox);
    }).then(function (contact) {
        console.log("Contact info for: " + mailbox)
        console.log("    " + contact.DisplayName);
        console.log("    " + contact.Email);
        var i = 0;
        for (i = 0; i < contact.phoneNumbers.length; i++) {
            console.log("    " + contact.phoneNumbers[i]);
        }
    }).catch(function (error) {
        console.log("Error!", error);
    });
