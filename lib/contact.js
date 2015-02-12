require('./msexchange.js');
require('es6-promise').polyfill(); // Promise polyfill global replacement
var xml2js = require('xml2js');

function contactInfo(displayName, email, phoneNumbers) {
    this.DisplayName = displayName;
    this.Email = email;
    this.phoneNumbers = phoneNumbers;
}
function contacts(msExchange, email) {
    this._session = msExchange;
}

contacts.prototype.getDetails = function(email) {
    console.log("Getting contact info for " + email);
    var soapReq = this._session.createEwsRequest(
        '<soap:Body>' +
        '<m:ResolveNames ReturnFullContactData="true" SearchScope="ContactsActiveDirectory">' +
        '<m:UnresolvedEntry>'+email+'</m:UnresolvedEntry>' +
        '</m:ResolveNames>' +
        '</soap:Body>');
    return this._session.makeEwsRequest(soapReq).then(function (response) {
        return new Promise(function (resolve, reject) {
            if (response == null) {
                reject(new Error("Contact couldnt be found"));
            }
            xml2js.parseString(response, function (err, xmlObj) {
                var contactxml = xmlObj['s:Envelope']['s:Body'][0]['m:ResolveNamesResponse'][0]['m:ResponseMessages'][0]
                                    ['m:ResolveNamesResponseMessage'][0]['m:ResolutionSet'][0]['t:Resolution'][0]
                                    ['t:Contact'][0];
                var emailAddress = xmlObj['s:Envelope']['s:Body'][0]['m:ResolveNamesResponse'][0]['m:ResponseMessages'][0]
                                    ['m:ResolveNamesResponseMessage'][0]['m:ResolutionSet'][0]['t:Resolution'][0]
                                    ['t:Mailbox'][0]['t:EmailAddress'][0];
                var contact = null;
                if (contactxml != null) {
                    var i;
                    var phoneNumbers = [];
                    for (i = 0; i < contactxml['t:PhoneNumbers'].length; i++) {
                        phoneNumbers.push(contactxml['t:PhoneNumbers'][i]['t:Entry'][0]._);
                    }
                    contact = new contactInfo(contactxml['t:DisplayName'][0], emailAddress, phoneNumbers);
                }
                resolve(contact);
            });
        });
    });
} 

contacts.prototype.getRoomLists = function () {
    //https://msdn.microsoft.com/en-us/library/office/hh532566(v=exchg.80).aspx
    console.log("Getting room lists");
    var soapReq = this._session.createEwsRequest('<soap:Body><m:GetRoomLists /></soap:Body>');
    return this._session.makeEwsRequest(soapReq)
    .then(function (response) {
        return new Promise(function (resolve, reject) {
            xml2js.parseString(response, function (err, xmlObj) {
                var roomLists = xmlObj['s:Envelope']['s:Body'][0]
                                    ['GetRoomListsResponse'][0]['m:RoomLists'][0]['t:Address'];
                var addresses = null;
                if (roomLists != null) {
                    addresses = roomLists
                        .map(function (obj) {
                            return new roomList(obj['Name'][0], obj['EmailAddress'][0]);
                        });
                }
                resolve(addresses);
            });
        });
    });
}

contacts.prototype.getRooms = function (roomListAddress) {
    //https://msdn.microsoft.com/en-us/library/office/hh532566(v=exchg.80).aspx
    console.log("Getting room lists");
    var soapReq = this._session.createEwsRequest('<soap:Body><m:ExpandDL><m:Mailbox><t:EmailAddress>' + roomListAddress + '</t:EmailAddress></m:Mailbox></m:ExpandDL></soap:Body>');
    return this._session.makeEwsRequest(soapReq)
    .then(function (response) {
        return new Promise(function (resolve, reject) {
            xml2js.parseString(response, function (err, xmlObj) {
                var rooms = xmlObj['s:Envelope']['s:Body'][0]
                                    ['m:ExpandDLResponse'][0]['m:ResponseMessages'][0]['m:ExpandDLResponseMessage'][0]['m:DLExpansion'];
                var roomAddresses = null;
                if (rooms != null) {
                    roomAddresses = rooms[0]['t:Mailbox']
                        .map(function (obj) {
                            return new roomResource(obj['Name'][0], obj['EmailAddress'][0]);
                        });
                }
                resolve(addresses);
            });
        });
    });
}

exports.Contacts = contacts;
