require('./../ews.js');
require('es6-promise').polyfill(); // Promise polyfill global replacement
var xml2js = require('xml2js');

function ContactInfo(displayName, email, phoneNumbers) {
    this.DisplayName = displayName;
    this.Email = email;
    this.phoneNumbers = phoneNumbers;
}
function RoomList(name, emailAddress) {
    this.name = name;
    this.emailAddress = emailAddress;
}
function RoomResource(name, emailAddress) {
    this.name = name;
    this.emailAddress = emailAddress;
}

function Contacts(msExchange, email) {
    this._session = msExchange;
}

Contacts.prototype.getDetails = function (email) {
    var soapReq = this._session.createEwsRequest(
        '<soap:Body>' +
        '<m:ResolveNames ReturnFullContactData="true" SearchScope="ContactsActiveDirectory">' +
        '<m:UnresolvedEntry>'+email+'</m:UnresolvedEntry>' +
        '</m:ResolveNames>' +
        '</soap:Body>');
    return this._session.makeEwsRequest(soapReq).then(function (response) {
        return new Promise(function (resolve, reject) {
            if (response == null) {
                reject(new Error(email + " couldn't be found in contacts."));
            }
            xml2js.parseString(response, { explicitArray: false }, function (err, xmlObj) {
                var contact = xmlObj['s:Envelope']['s:Body']['m:ResolveNamesResponse']
                                    ['m:ResponseMessages']['m:ResolveNamesResponseMessage']
                                    ['m:ResolutionSet']['t:Resolution']['t:Contact'];
                var mailbox = xmlObj['s:Envelope']['s:Body']
                                    ['m:ResolveNamesResponse']['m:ResponseMessages']
                                    ['m:ResolveNamesResponseMessage']['m:ResolutionSet']
                                    ['t:Resolution']['t:Mailbox'];
                var contactInfo = null;
                if (contact != null) {
                    var i;
                    var phoneNumbers = [];
                    if (contact['t:PhoneNumbers'].constructor === Array) {
                        for (i = 0; i < contact['t:PhoneNumbers'].length; i++) {
                            phoneNumbers.push(contact['t:PhoneNumbers'][i]['t:Entry']._);
                        }
                    } else phoneNumbers.push(contact['t:PhoneNumbers']['t:Entry']._);
                    contactInfo = new ContactInfo(contact['t:DisplayName'], mailbox['t:EmailAddress'], phoneNumbers);
                }
                resolve(contactInfo);
            });
        });
    });
} 

Contacts.prototype.getRoomLists = function () {
    //https://msdn.microsoft.com/en-us/library/office/hh532566(v=exchg.80).aspx
    var soapReq = this._session.createEwsRequest('<soap:Body><m:GetRoomLists /></soap:Body>');
    return this._session.makeEwsRequest(soapReq)
    .then(function (response) {
        return new Promise(function (resolve, reject) {
            xml2js.parseString(response, { explicitArray: false }, function (err, xmlObj) {
                var roomLists = xmlObj['s:Envelope']['s:Body']
                                      ['GetRoomListsResponse']
                                      ['m:RoomLists']['t:Address'];
                var addresses = null;
                if (roomLists != null) {
                    addresses = roomLists
                        .map(function (obj) {
                            return new RoomList(obj['Name'], obj['EmailAddress']);
                        });
                }
                resolve(addresses);
            });
        });
    });
}

Contacts.prototype.getRooms = function (roomListAddress) {
    //https://msdn.microsoft.com/en-us/library/office/hh532566(v=exchg.80).aspx
    var soapReq = this._session.createEwsRequest(
           '<soap:Body><m:ExpandDL><m:Mailbox>'+
              '<t:EmailAddress>' + roomListAddress + '</t:EmailAddress>'+
           '</m:Mailbox></m:ExpandDL></soap:Body>'
           );
    return this._session.makeEwsRequest(soapReq)
    .then(function (response) {
        return new Promise(function (resolve, reject) {
            xml2js.parseString(response, { explicitArray: false }, function (err, xmlObj) {
                var rooms = xmlObj['s:Envelope']['s:Body']
                                  ['m:ExpandDLResponse']['m:ResponseMessages']
                                  ['m:ExpandDLResponseMessage']['m:DLExpansion'];
                var roomAddresses = null;
                if (rooms != null) {
                    roomAddresses = rooms['t:Mailbox']
                        .map(function (obj) {
                            return new RoomResource(obj['Name'], obj['EmailAddress']);
                        });
                }
                resolve(addresses);
            });
        });
    });
}

exports.Contacts = Contacts;
