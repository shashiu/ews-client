'use strict';
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
                    // Phone numbers come in as follows
                    // <t:PhoneNumbers>
                    //  <t:Entry Key="AssistantPhone"></t:Entry>
                    //  <t:Entry Key="BusinessFax">425-895-4750</t:Entry>
                    //  <t:Entry Key="BusinessPhone"> +1 425 895 4754</t:Entry>
                    // </t:PhoneNumbers>
                    var phoneNumbers = {};
                    if (contact['t:PhoneNumbers']['t:Entry'] != null &&
                        contact['t:PhoneNumbers']['t:Entry'].constructor === Array) {
                        for (i = 0; i < contact['t:PhoneNumbers']['t:Entry'].length; i++) {
                            if (contact['t:PhoneNumbers']['t:Entry'][i]._ != undefined) {
                                phoneNumbers[contact['t:PhoneNumbers']['t:Entry'][i].$.Key] =
                                             contact['t:PhoneNumbers']['t:Entry'][i]._;
                            }
                        }
                    } else {
                        phoneNumbers[contact['t:PhoneNumbers']['t:Entry'].$.Key] =
                                     contact['t:PhoneNumbers']['t:Entry']._;
                    }
                    contactInfo = new ContactInfo(contact['t:DisplayName'],mailbox['t:EmailAddress'], phoneNumbers);
                }

                resolve(contactInfo);
            });
        });
    });
} 

/*
 A note about address books - these are available for use via the outlook APIs and not from exchange.
 Exchnage 2010 has the notion of 'Offline Address Book' (URL in the response to the autodiscover API),
 which are loaded by outlook and kept in sync in the background (using BITS).
 
 It is not possible to enumerate rooms using the EWS API (service.FindItems method or any other
 method). The WellknownFolder.Contacts points to the user's personal contacts folder and not the GAL.
 What outlook displays as the GAL either comes from AD or from the local offline address book.
 The MFCMAPI tool shows they have a PID_DISPLAY_TYPE_EX property of 7 (DT_MAILBOX_USER | DT_ROOM).
*/
Contacts.prototype.enumerateAllRooms = function () {
    // NOT WORKING 
    var ewsSoapReq = this._session.createEwsRequest(
        '<soap:Body>' +
          '<m:FindItem Traversal="Shallow"><m:ItemShape><t:BaseShape>AllProperties</t:BaseShape></m:ItemShape>' +
            '<m:IndexedPageItemView MaxEntriesReturned="5" Offset="0" BasePoint="Beginning" />' +
            //'<m:Restriction><t:IsEqualTo><t:FieldURI PropertyTag="14597" PropertyType="Long" />' +
            '<m:Restriction><t:IsEqualTo><t:ExtendedFieldURI PropertyTag="14597" PropertyType="Integer" />' +
              '<t:FieldURIOrConstant><t:Constant Value="7" /></t:FieldURIOrConstant>' +
            '</t:IsEqualTo></m:Restriction>' +
            '<m:ParentFolderIds><t:DistinguishedFolderId Id="contacts" /></m:ParentFolderIds>' +
          '</m:FindItem>' +
        '</soap:Body>');
    return this._session.makeEwsRequest(ewsSoapReq).then(function (response) {
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
