/*
 * A library to establish an authenticated session with MS Exchange 
  * and make EWS requests to get to a mailbox's calendar and contacts
 */
'use strict';
var ntlm = require('httpntlm').ntlm;
var httpreq = require('request');
var keepalive = require('keep-alive-agent').Secure;
var xml2js = require('xml2js');
var soap = require('soap');
require('es6-promise').polyfill(); // Promise polyfill global replacement
require('./../ews.js');

function ewsSoapRequest(body, headers, options) {
    this.body = body;
    this.headers = headers;
    this.options = options;
}

function MSExchange(user, pass, domain) {
    this.options = {
        url: null,
        username: user,
        password: pass,
        workstation: '',
        domain: domain
    };
}

MSExchange.prototype.setSvcUrl = function (url) {
    this.options.url = url;
}
MSExchange.prototype.createEwsRequest = function (soapBody) {
    var soapReq = '<?xml version="1.0" encoding="utf-8"?>' +
        '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '+
            'xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" '+
            'xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" ' +
            'xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
            '<soap:Header>' +
               '<t:RequestServerVersion Version="Exchange2010_SP1" />' +
               '<t:TimeZoneContext>' + '<t:TimeZoneDefinition Id="Pacific Standard Time" />' + '</t:TimeZoneContext>' +
            '</soap:Header>' +
            soapBody +
        '</soap:Envelope>';
    var soapHeader = {
        'Content-Type': 'text/xml; charset=utf-8', // required
        'Accept': 'text/xml',
        'Connection': 'keep-alive',
    };
    return new ewsSoapRequest(soapReq, soapHeader, this.options);
}
/*
 * Helper function to do an HTTP Post with NTLM auth
 * returns a Promise object that can be chained.
 */
MSExchange.prototype.makeEwsRequest = function (request) {
    return new Promise(function (resolve, reject) {
        /*
        * NTLM auth works on a session basis, we have to keep the TCP session alive
        * across the 2 NTLM auth messages
        */
        var kaAgent = new keepalive();

        /* 2 stage NTLM authentication */
        var type1msg = ntlm.createType1Message(request.options);
        if (request.headers == null) {
            request.headers = {
                'Content-Type': 'text/xml; charset=utf-8', // required
                'Accept': 'text/xml',
                'Connection': 'keep-alive'
            };
        }
        request.headers['Authorization'] = type1msg;
        httpreq.post(
            {
                uri: request.options.url,
                headers: request.headers,
                body: request.body,
                agent: kaAgent,
                strictSSL: false,
            },
            function (err, res) {
                if (err) {
                    console.log(err);
                    reject(err);
                    return;
                }
                if (res.statusCode == 401) {
                    if (!res.headers['www-authenticate'])
                        return console.log('www-authenticate not found on response of second request');

                    var type2msg = ntlm.parseType2Message(res.headers['www-authenticate']);
                    var type3msg = ntlm.createType3Message(type2msg, request.options);

                    /* 
                     * Close this HTTP connection after this request ?
                     * Perhaps keeping it running might offer better performance
                     */
                    request.headers['Connection'] = 'Close';
                    request.headers['Authorization'] = type3msg;

                    httpreq.post({
                        uri: request.options.url,
                        headers: request.headers,
                        body: request.body,
                        allowRedirects: false,
                        agent: kaAgent,
                        strictSSL: false
                    }, function (err, res2) {
                        if (err) {
                            console.log("Error:" + err);
                            reject(err);
                        } else {
                            resolve(res2.body);
                        }
                    });

                } else {
                    if (res.statusCode == 200)
                        resolve(res.body);
                    else
                        reject(new Error("Error: received error code " + res.statusCode.toString()));
                }
            });
    });
}

MSExchange.prototype.autoDiscover = function (mailbox) {

    var realm = "https://autodiscover." + mailbox.split("@")[1];
    this.options.url = realm + "/autodiscover/autodiscover.xml";

    var req = new ewsSoapRequest(
        '<Autodiscover xmlns="http://schemas.microsoft.com/exchange/autodiscover/outlook/requestschema/2006">' +
            '<Request>' +
                '<EMailAddress>' + mailbox + '</EMailAddress>' +
                '<AcceptableResponseSchema>http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a</AcceptableResponseSchema>' +
            '</Request>' +
        '</Autodiscover>',
        null, this.options);

    return this.makeEwsRequest(req)
    .then(function (response) {
        //console.log(response);
        return new Promise(function (resolve, reject) {
            xml2js.parseString(response, { explicitArray: false }, function (err, xmlObj) {
                if (xmlObj == null || xmlObj.Autodiscover == null || xmlObj.Autodiscover.Response.Error != null) {
                    var errStr = "Autodiscover Error: " + response;
                    if (xmlObj != null && xmlObj.Autodiscover != null) {
                        /* well formed error string */
                        errStr = "Autodiscover Error: " + xmlObj.Autodiscover.Response.Error.Message;
                    }
                    reject(new Error(errStr));
                } else {
                    /* no errors */
                    var i = 0;
                    for (i = 0; i < xmlObj.Autodiscover.Response.Account.Protocol.length; i++) {
                        if (xmlObj.Autodiscover.Response.Account.Protocol[i].Type == "EXPR") {
                            /* use the external URL to access the service */
                            this.setSvcUrl(xmlObj.Autodiscover.Response.Account.Protocol[i].EwsUrl);
                        }
                    }
                    resolve();
                }
            }.bind(this));
        }.bind(this));
    }.bind(this));
}
/*
 * Get a well known folder:
 * 'Calendar', 'Inbox', 'Contacts'
 */
MSExchange.prototype.getFolder = function (wellknownFoldername) {
    var ewsSoapReq = this.createEwsRequest(
              '<soap:Body><GetFolder xmlns="http://schemas.microsoft.com/exchange/services/2006/messages" '+
                   'xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"> '+
                '<FolderShape><t:BaseShape>Default</t:BaseShape></FolderShape>'+
                '<FolderIds><t:DistinguishedFolderId Id="' + wellknownFoldername + '"/></FolderIds>'+
              '</GetFolder></soap:Body>'
            );

    return this.makeEwsRequest(ewsSoapReq)
    .then(function (response) {
        return new Promise(function (resolve, reject) {
            xml2js.parseString(response, { explicitArray: false }, function (err, xmlObj) {
                resolve(
                    {
                        FolderId: xmlObj['s:Envelope']['s:Body']['m:GetFolderResponse']
                            ['m:ResponseMessages']['m:GetFolderResponseMessage']['m:Folders']
                            ['t:CalendarFolder']['t:FolderId'].$.Id,
                        ChangeKey: xmlObj['s:Envelope']['s:Body']['m:GetFolderResponse']
                            ['m:ResponseMessages']['m:GetFolderResponseMessage']['m:Folders']
                            ['t:CalendarFolder']['t:FolderId'].$.ChangeKey
                    }
                ); // return object with folder id and chnageKey for folder
            });
        });
    });
}

exports.MSExchange = MSExchange;
