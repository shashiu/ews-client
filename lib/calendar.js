require('./../ews.js');
require('es6-promise').polyfill(); // Promise polyfill global replacement
var xml2js = require('xml2js');

function CalendarItem(
            id,
            changeKey
            ) {
    this.Id = id;
    this.ChangeKey = changeKey;
    this.Attendees = null;
    this.Subject = null;
    this.Start = null;
    this.End = null;
    this.MeetingUrl = null;
    this.UCConferenceSetting = null;
    this.MeetingAudioOptions = null;
    this.Body = null;
}

function Calendar(msExchangeSession, mailbox, calendarId, changeKey) {
    this._session = msExchangeSession;
    this._mailbox = mailbox;
    this._calendarId = calendarId;
    this._changeKey = changeKey;
}

_parseEntry = function (obj) {
    var ucConferenceSetting = null;
    var meetingUrl = null;
    var requiredAttendees = null;
    var audioOptions = null;

    if (obj['t:ExtendedProperty'] != null) {
        // if the lync property exists.
        var i;
        for (i = 0; i < obj['t:ExtendedProperty'].length; i++) {
            if (obj['t:ExtendedProperty'][i]['t:ExtendedFieldURI'].$.PropertyName == 'OnlineMeetingExternalLink') {
                meetingUrl = obj['t:ExtendedProperty'][i]['t:Value'].toString();
            }
            if (obj['t:ExtendedProperty'][i]['t:ExtendedFieldURI'].$.PropertyName == 'UCConferenceSetting') {
                ucConferenceSetting = obj['t:ExtendedProperty'][i]['t:Value'].toString();
            }
            if (obj['t:ExtendedProperty'][i]['t:ExtendedFieldURI'].$.PropertyName == 'GTMAudioOptions') {
                audioOptions = obj['t:ExtendedProperty'][i]['t:Value'].toString();
            }
            if (obj['t:ExtendedProperty'][i]['t:ExtendedFieldURI'].$.PropertyName == 'RequiredAttendees') {
                requiredAttendees = obj['t:ExtendedProperty'][i]['t:Value'].toString();
            }
        }
    }
    var entry = new CalendarItem(obj['t:ItemId'].$.Id, obj['t:ItemId'].$.ChangeKey);

    if (obj['t:Body'] != null) {
        entry.Body = obj['t:Body'];
    }
    entry.Subject = obj['t:Subject'];
    entry.Start = obj['t:Start'];
    entry.End = obj['t:End'];
    entry.MeetingUrl = meetingUrl;
    entry.UCConferenceSetting = ucConferenceSetting;
    entry.MeetingAudioOptions = audioOptions;
    entry.Attendees = requiredAttendees;

    return entry;
}

Calendar.prototype.getItemDetails = function (item) {
    var ewsSoapReq = this._session.createEwsRequest(
            '<soap:Body><m:GetItem>'+
                '<m:ItemShape><t:BaseShape>AllProperties</t:BaseShape>' +
                   '<t:AdditionalProperties>'+
                        '<t:ExtendedFieldURI DistinguishedPropertySetId="PublicStrings" PropertyName="OnlineMeetingExternalLink" PropertyType="String" />' +
                        '<t:ExtendedFieldURI DistinguishedPropertySetId="PublicStrings" PropertyName="UCConferenceSetting" PropertyType="String" />' +
                        '<t:ExtendedFieldURI DistinguishedPropertySetId="PublicStrings" PropertyName="GTMAudioOptions" PropertyType="String" />' +
                        '<t:ExtendedFieldURI DistinguishedPropertySetId="PublicStrings" PropertyName="RequiredAttendees" PropertyType="String" />' +
                        '<t:ExtendedFieldURI DistinguishedPropertySetId="PublicStrings" PropertyName="OptionalAttendees" PropertyType="String" />' +
                        '</t:AdditionalProperties>' +
                '</m:ItemShape>'+
                '<m:ItemIds><t:ItemId Id="' + item.Id + '" ChangeKey="' + item.ChangeKey + '" /></m:ItemIds>' +
            '</m:GetItem></soap:Body>'
            );
    return this._session.makeEwsRequest(ewsSoapReq)
    .then(function (response) {
        return new Promise(function (resolve, reject) {
            xml2js.parseString(response, { explicitArray: false }, function (err, xmlObj) { // ignoreAttrs : true ??
                var entry = xmlObj['s:Envelope']['s:Body']['m:GetItemResponse']
                                    ['m:ResponseMessages']['m:GetItemResponseMessage']
                                    ['t:Items']['t:CalendarItem']
                resolve(_parseEntry(entry)); // return calendar entry array
            });
        });
    });
}

Calendar.prototype.getEntries = function (startTime, endTime, maxEntries) {
    if (maxEntries == null) maxEntries = 5;
    if (startTime == null) startTime =  (new Date()).toISOString();
    if (endTime == null) {
        endDate = new Date();
        endDate.setHours(endDate.getHours() + 24);
        endTime = endDate.toISOString();
    }
    var parentFolderIds = null;
    if (this._calendarId != null && this._changeKey != null) {
        parentFolderIds = '<m:ParentFolderIds><t:FolderId Id="' + this._calendarId +
                            '" ChangeKey="' + this._changeKey + '" /></m:ParentFolderIds>';
    } else {
        parentFolderIds = '<m:ParentFolderIds><t:DistinguishedFolderId Id="calendar"><t:Mailbox>' +
                            '<t:EmailAddress>' + this._mailbox + '</t:EmailAddress>' +
                          '</t:Mailbox></t:DistinguishedFolderId></m:ParentFolderIds>';
    }
    var ewsSoapReq = this._session.createEwsRequest(
            '<soap:Body><m:FindItem Traversal="Shallow">' +
              '<m:ItemShape><t:BaseShape>IdOnly</t:BaseShape>' +
                '<t:AdditionalProperties>' +
                  '<t:FieldURI FieldURI="item:Subject" />' +
                  '<t:FieldURI FieldURI="calendar:Start" />' +
                  '<t:FieldURI FieldURI="calendar:End" />' +
                  '<t:ExtendedFieldURI DistinguishedPropertySetId="PublicStrings" PropertyName="OnlineMeetingExternalLink" PropertyType="String" />' +
                  '<t:ExtendedFieldURI DistinguishedPropertySetId="PublicStrings" PropertyName="UCConferenceSetting" PropertyType="String" />' +
                  '<t:ExtendedFieldURI DistinguishedPropertySetId="PublicStrings" PropertyName="GTMAudioOptions" PropertyType="String" />' +
                  '<t:ExtendedFieldURI DistinguishedPropertySetId="PublicStrings" PropertyName="RequiredAttendees" PropertyType="String" />' +
                  '<t:ExtendedFieldURI DistinguishedPropertySetId="PublicStrings" PropertyName="OptionalAttendees" PropertyType="String" />' +
                '</t:AdditionalProperties>'+
              '</m:ItemShape>' +
              '<m:CalendarView MaxEntriesReturned="'+ maxEntries + '" StartDate="' + startTime + '" EndDate="' + endTime + '"/>' +
              parentFolderIds +
            '</m:FindItem></soap:Body>');

    return this._session.makeEwsRequest(ewsSoapReq)
    .then(function (response) {
        return new Promise(function (resolve, reject) {
            var entries = null;
            xml2js.parseString(response, { explicitArray: false }, function (err, xmlObj) {
                var items = xmlObj['s:Envelope']['s:Body']['m:FindItemResponse']
                                  ['m:ResponseMessages']['m:FindItemResponseMessage']
                                  ['m:RootFolder']['t:Items']['t:CalendarItem'];
                if (items != null) {
                    if (items.constructor !== Array) items = [items];
                    entries = items.map(function (obj) {
                            return _parseEntry(obj);
                          });
                } 
                resolve(entries); // return calendar entry array
            });
        });
    });
}

Calendar.prototype.getAvailabilityForRoom = function (roomMailbox, startTime, endTime) {

    if (startTime == null) startTime = (new Date()).toISOString();
    if (endTime == null) { //+ 24 hrs
        endDate = new Date();
        endDate.setHours(endDate.getHours() + 24);
        endTime = endDate.toISOString();
    }

    var ewsSoapReq = this._session.createEwsRequest(
                        '<soap:Body><m:GetUserAvailabilityRequest>'+
                           '<m:MailboxDataArray><t:MailboxData>' +
                             '<t:Email><t:Address>' + roomMailbox + '</t:Address></t:Email>' +
                             '<t:AttendeeType>Room</t:AttendeeType>' +
                             '<t:ExcludeConflicts>false</t:ExcludeConflicts>'+
                           '</t:MailboxData></m:MailboxDataArray>' +
                           '<t:FreeBusyViewOptions><t:TimeWindow><t:StartTime>' + startTime + '</t:StartTime>' +
                               '<t:EndTime>' + endTime + '</t:EndTime></t:TimeWindow>'+
                               '<t:MergedFreeBusyIntervalInMinutes>30</t:MergedFreeBusyIntervalInMinutes>' +
                               '<t:RequestedView>Detailed</t:RequestedView>'+
                           '</t:FreeBusyViewOptions>'+
                        '</m:GetUserAvailabilityRequest></soap:Body>'
                        );

    return this._session.makeEwsRequest(ewsSoapReq)
    .then(function (response) {
        return new Promise(function (resolve, reject) {

            xml2js.parseString(response, { explicitArray: false }, function (err, xmlObj) {
                var response = xmlObj['s:Envelope']['s:Body']
                                ['GetUserAvailabilityResponse']['FreeBusyResponseArray']
                                ['FreeBusyResponse'];
                if (response['ResponseMessage']['ResponseCode'] != 'NoError') {
                    reject(new Error(response['ResponseMessage']['ResponseCode']));
                    return;
                }

                var calendarEvents = response['FreeBusyView']['CalendarEventArray'];
                var entries = null;
                if (calendarEvents != null) {
                    calEvents = calendarEvents['CalendarEvent'];
                    if (calEvents.constructor !== Array) calEvents = [calEvents];
                    entries = calEvents.map(function (obj) {
                            var item = new CalendarItem(
                                obj['CalendarEventDetails']['ID'],
                                obj['CalendarEventDetails']['ChangeKey']
                            );
                            item.Subject = obj['CalendarEventDetails']['Subject'];
                            item.Start = obj['StartTime'];
                            item.End = obj['EndTime'];
                            return item;
                        });
                }

                resolve(entries); 
            });
        });
    });
}



exports.Calendar = Calendar;
