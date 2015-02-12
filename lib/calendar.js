require('./msexchange.js');
require('es6-promise').polyfill(); // Promise polyfill global replacement
var xml2js = require('xml2js');

function calendarItem(subject, start, end, id, changeKey, meetingUrl, conferenceSettings, audioOptions) {
    this.Attendees = null;
    this.Subject = subject;
    this.Start = start;
    this.End = end;
    this.Id = id;
    this.ChangeKey = changeKey;
    this.MeetingUrl = meetingUrl;
    this.UCConferenceSetting = conferenceSettings;
    this.GTMAudioOptions = audioOptions;
}
function roomList(name, emailAddress) {
    this.name = name;
    this.emailAddress = emailAddress;
}
function roomResource(name, emailAddress) {
    this.name = name;
    this.emailAddress = emailAddress;
}

/* 
  Calendar Constructor
 */
function calendar(msExchangeSession) {
    this._session = msExchangeSession;
}

calendar.prototype.getCalendar = function (calendarId, changeKey, startTime, endTime) {

    if (startTime == null) {
        startDate = new Date();
        startTime = startDate.toISOString();
    }

    if (endTime == null) {
        endDate = new Date();
        endDate.setHours(endDate.getHours() + 24);
        endTime = endDate.toISOString();
    }

    var ewsSoapReq = this._session.createEwsRequest(
           '<soap:Body><m:FindItem Traversal="Shallow"><m:ItemShape><t:BaseShape>IdOnly</t:BaseShape><t:AdditionalProperties>' +
            '<t:FieldURI FieldURI="item:Subject" /><t:FieldURI FieldURI="calendar:Start" />' +
            '<t:FieldURI FieldURI="calendar:End" /></t:AdditionalProperties>' +
               '</m:ItemShape><m:CalendarView MaxEntriesReturned="5" StartDate="' + startTime + '" EndDate="' + endTime + '" /><m:ParentFolderIds>' +
                '<t:FolderId Id="' + calendarId + '" ChangeKey="' + changeKey + '" /></m:ParentFolderIds>' +
             '</m:FindItem></soap:Body>');

    
    return this._session.makeEwsRequest(ewsSoapReq)
    .then(function (response) {
        //console.log(response);
        return new Promise(function (resolve, reject) {
            xml2js.parseString(response, function (err, xmlObj) {
                var entries = xmlObj['s:Envelope']['s:Body'][0]['m:FindItemResponse'][0]
                                    ['m:ResponseMessages'][0]['m:FindItemResponseMessage'][0]['m:RootFolder'][0]
                                    ['t:Items'][0]['t:CalendarItem']
                    .map(function(obj) {
                        return new calendarItem(obj['t:Subject'][0], obj['t:Start'][0], obj['t:End'][0]);
                    });
                resolve(entries); // return calendar entry array
            });
        });
    });
}

calendar.prototype.getItem = function (itemId, changeKey) {
    var ewsSoapReq = this._session.createEwsRequest(
            '<soap:Body><m:GetItem><m:ItemShape><t:BaseShape>AllProperties</t:BaseShape>' +
            '<t:AdditionalProperties><t:ExtendedFieldURI DistinguishedPropertySetId="PublicStrings" PropertyName="OnlineMeetingExternalLink" PropertyType="String" /></t:AdditionalProperties>' +
            '</m:ItemShape><m:ItemIds><t:ItemId Id="' + itemId + '" ChangeKey="' + changeKey + '" /></m:ItemIds></m:GetItem></soap:Body>'
            );

    return this._session.makeEwsRequest(ewsSoapReq)
    .then(function (response) {
        //console.log(response);
        return new Promise(function (resolve, reject) {
            xml2js.parseString(response, function (err, xmlObj) {
                var entries = xmlObj['s:Envelope']['s:Body'][0]['m:GetItemResponse'][0]
                                    ['m:ResponseMessages'][0]['m:GetItemResponseMessage'][0]
                                    ['t:Items'][0]['t:CalendarItem']
                    .map(function (obj) {
                        var meetingUrl = null;
                        if (obj['t:ExtendedProperty'] != null)
                            meetingUrl = ['t:ExtendedProperty'][0]['t:Value'];
                        return new calendarItem(obj['t:Subject'][0], obj['t:Start'][0], obj['t:End'][0], obj['t:ItemId'][0]['Id'], meetingUrl);
                    });
                resolve(entries); // return calendar entry array
            });
        });
    });
}

calendar.prototype.getEntries = function (mailbox, startTime, endTime) {

    if (startTime == null) {
        startDate = new Date();
        startTime = startDate.toISOString();
    }

    if (endTime == null) {
        endDate = new Date();
        endDate.setHours(endDate.getHours() + 24);
        endTime = endDate.toISOString();
    }

    var ewsSoapReq = this._session.createEwsRequest(
            //'<soap:Body><m:FindItem Traversal="Shallow"><m:ItemShape><t:BaseShape>IdOnly</t:BaseShape>'+
            '<soap:Body><m:FindItem Traversal="Shallow">' +
            '<m:ItemShape><t:BaseShape>AllProperties</t:BaseShape>' +
            '<t:AdditionalProperties>' +
              '<t:FieldURI FieldURI="item:Subject" />' +
              '<t:FieldURI FieldURI="calendar:Start" />' +
              '<t:FieldURI FieldURI="calendar:End" />' +
              //'<t:ExtendedFieldURI DistinguishedPropertySetId="PublicStrings" PropertyName="UCOpenedConferenceID" PropertyType="String" />' +
              '<t:ExtendedFieldURI DistinguishedPropertySetId="PublicStrings" PropertyName="OnlineMeetingExternalLink" PropertyType="String" />' +
              '<t:ExtendedFieldURI DistinguishedPropertySetId="PublicStrings" PropertyName="UCConferenceSetting" PropertyType="String" />' +
              '<t:ExtendedFieldURI DistinguishedPropertySetId="PublicStrings" PropertyName="GTMAudioOptions" PropertyType="String" />' +
              '<t:ExtendedFieldURI DistinguishedPropertySetId="PublicStrings" PropertyName="RequiredAttendees" PropertyType="String" />' +
              '<t:ExtendedFieldURI DistinguishedPropertySetId="PublicStrings" PropertyName="OptionalAttendees" PropertyType="String" />' +
            '</t:AdditionalProperties>'+
            '</m:ItemShape>' +
            '<m:CalendarView MaxEntriesReturned="1000" StartDate="'+startTime+'" EndDate="'+endTime+'"/>'+
            '<m:ParentFolderIds><t:DistinguishedFolderId Id="calendar"><t:Mailbox>'+
                '<t:EmailAddress>' + mailbox + '</t:EmailAddress>' +
                '</t:Mailbox></t:DistinguishedFolderId></m:ParentFolderIds></m:FindItem></soap:Body>');

    return this._session.makeEwsRequest(ewsSoapReq)
    .then(function (response) {
        //console.log(response);
        return new Promise(function (resolve, reject) {
            var entries = null;
            //console.log(response);
            xml2js.parseString(response, function (err, xmlObj) {
                var items = xmlObj['s:Envelope']['s:Body'][0]['m:FindItemResponse'][0]
                                    ['m:ResponseMessages'][0]['m:FindItemResponseMessage'][0]['m:RootFolder'][0]
                                    ['t:Items'][0]['t:CalendarItem']
                if (items != null) {
                    entries = items.map(function (obj) {
                            var ucConferenceSetting = null;
                            var meetingUrl = null;
                            var requiredAttendees = null;
                            var audioOptions = null;
                            if (obj['t:ExtendedProperty'] != null) {
                                // if the lync property exists.
                                var i;
                                for (i = 0; i < obj['t:ExtendedProperty'].length; i++) {
                                    if (obj['t:ExtendedProperty'][i]['t:ExtendedFieldURI'][0].$.PropertyName == 'OnlineMeetingExternalLink') {
                                        meetingUrl = obj['t:ExtendedProperty'][i]['t:Value'].toString();
                                    }
                                    if (obj['t:ExtendedProperty'][i]['t:ExtendedFieldURI'][0].$.PropertyName == 'UCConferenceSetting') {
                                        ucConferenceSetting = obj['t:ExtendedProperty'][i]['t:Value'].toString();
                                    }
                                    if (obj['t:ExtendedProperty'][i]['t:ExtendedFieldURI'][0].$.PropertyName == 'GTMAudioOptions') {
                                        audioOptions = obj['t:ExtendedProperty'][i]['t:Value'].toString();
                                    }
                                    if (obj['t:ExtendedProperty'][i]['t:ExtendedFieldURI'][0].$.PropertyName == 'RequiredAttendees') {
                                        requiredAttendees = obj['t:ExtendedProperty'][i]['t:Value'].toString();
                                    }
                                }
                            }
                            var entry = new calendarItem(
                                obj['t:Subject'][0], obj['t:Start'][0], obj['t:End'][0],
                                obj['t:ItemId'][0]['Id'], obj['t:ItemId'][0]['ChangeKey'],
                                meetingUrl,
                                ucConferenceSetting,
                                audioOptions
                            );
                        entry.Attendees = requiredAttendees;
                        return entry;
                    });
                } 
                resolve(entries); // return calendar entry array
            });
        });
    });
}

calendar.prototype.getAvailabilityForRoom = function (roomMailbox, startTime, endTime) {

    if (startTime == null) { // now
        startDate = new Date();
        startTime = startDate.toISOString();
    }

    if (endTime == null) { //+ 24 hrs
        endDate = new Date();
        endDate.setHours(endDate.getHours() + 24);
        endTime = endDate.toISOString();
    }

    console.log("Getting calendar entries for " + roomMailbox + ', from ' + startTime + ', to ' + endTime)
    var ewsSoapReq = this._session.createEwsRequest(
                        '<soap:Body><m:GetUserAvailabilityRequest><m:MailboxDataArray><t:MailboxData>' +
                        '<t:Email><t:Address>' + roomMailbox + '</t:Address></t:Email>' +
                        '<t:AttendeeType>Room</t:AttendeeType><t:ExcludeConflicts>false</t:ExcludeConflicts></t:MailboxData></m:MailboxDataArray>' +
                        '<t:FreeBusyViewOptions><t:TimeWindow><t:StartTime>' + startTime + '</t:StartTime><t:EndTime>' + endTime + '</t:EndTime></t:TimeWindow><t:MergedFreeBusyIntervalInMinutes>30</t:MergedFreeBusyIntervalInMinutes>' +
                        '<t:RequestedView>Detailed</t:RequestedView></t:FreeBusyViewOptions></m:GetUserAvailabilityRequest>' +
                        '</soap:Body>');
    return this._session.makeEwsRequest(ewsSoapReq)
    .then(function (response) {
        //console.log(response);
        return new Promise(function (resolve, reject) {
            xml2js.parseString(response, function (err, xmlObj) {
                var calendarEvents = xmlObj['s:Envelope']['s:Body'][0]
                                ['GetUserAvailabilityResponse'][0]
                                ['FreeBusyResponseArray'][0]['FreeBusyResponse'][0]['FreeBusyView'][0]
                                ['CalendarEventArray'];
                var entries = null;
                if (calendarEvents != null) {
                    entries = calendarEvents[0]['CalendarEvent']
                        .map(function(obj) {
                            //return new calendarEntry('subject', 'start', 'time');
                            return new calendarItem(
                                obj['CalendarEventDetails'][0]['Subject'][0],
                                obj['StartTime'][0],
                                obj['EndTime'][0],
                                obj['CalendarEventDetails'][0]['ID'][0]);
                        });
                }
                resolve(entries); // return calendar entry array
            });
        });
    });
}



exports.Calendar = calendar;
