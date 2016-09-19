/// <reference path="../../typings/index.d.ts" />

exports.LoadJsom = function (serverAbsoluteUrl) {
    var window = global;

    window.XMLHttpRequest = require("xmlhttprequest").XMLHttpRequest;
    const url = require("URL");
    var urlObject = url.parse(serverAbsoluteUrl, true, true);

    var relativeUrl = urlObject.pathname;
   
    if(urlObject.search){
        relativeUrl += urlObject.search;
    }

    if(urlObject.hash){
        relativeUrl += hash;
    }

    window._spPageContextInfo = {
        webAbsoluteUrl: serverAbsoluteUrl,
        webServerRelativeUrl: relativeUrl,
    }
    
    window.navigator = {
        userAgent: "Node"
    }

    window.formdigest = {
        value: '',
        tagName: 'INPUT',
        type: 'hidden'
    };

    window.location = urlObject;

    window.document = {
        URL: window.location.href,
        cookie: "",
        documentElement: {},
        getElementsByName: function (name) {
            if (name == '__REQUESTDIGEST') {
                return [window.formdigest];
            }
        },
        getElementsByTagName: function (name) {
            return [];
        }
    };

    //spjcsom

    window.escapeUrlForCallback = escapeUrlForCallback;
}



