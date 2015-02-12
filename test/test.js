var ews = require('./../ews');

// get web credentials from process environment
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

// Autodiscover URL to get to exchange web service
msExchange.setAuth(user, pass, domain);
console.log("Autodiscover");
msExchange.autoDiscover(mailbox)
.then(function () {

    console.log("Autodiscover Success!");
    _testCalendarAPIs(msExchange);
    _testContactAPIs(msExchange);

}).catch(function (error) {
    console.log("Error!", error);
});

_testCalendarAPIs = function (session) {
    console.log('Getting calendar entries for next 24 hours for session calendar');
    session.getFolder('calendar').then(function (folder) {
        var calendar = new ews.Calendar(session, null, folder.FolderId, folder.ChangeKey);
        calendar.getEntries().then(_displayCalendarEntries);
        console.log('Getting calendar entries for next 24 hours for mailbox ' + mailbox);
        var calendar2 = new ews.Calendar(session, mailbox);
        calendar2.getEntries().then(_displayCalendarEntries);
    });
}

_displayCalendarEntries = function (entries) {
    console.log('----------------------------------------------------------------------------');
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
}

_testContactAPIs = function (session) {
    var contacts = new ews.Contacts(session);
    console.log("Getting room lists");
    contacts.getRoomLists().then(function (roomLists) {
        if (roomLists == null) {
            console.log('    Room List is null (Exchange Admin hasn\'t configured them)');
        }
        console.log("Getting Contact details for " + mailbox);
        return contacts.getDetails(mailbox);
    }).then(function (contact) {
        console.log("    " + contact.DisplayName);
        console.log("    " + contact.Email);
        var i = 0;
        for (i = 0; i < contact.phoneNumbers.length; i++) {
            console.log("    " + contact.phoneNumbers[i]);
        }
    });
}