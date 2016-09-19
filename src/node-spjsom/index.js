exports.LoadJsom = function (serverAbsoluteUrl) {
    global.window = global;

    const vm = require('vm');
    window.XMLHttpRequest = require("xmlhttprequest").XMLHttpRequest;
    const url = require("URL");
    var urlObject = url.parse(serverAbsoluteUrl, true, true);

    var relativeUrl = urlObject.pathname;

    if (urlObject.search) {
        relativeUrl += urlObject.search;
    }

    if (urlObject.hash) {
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

    let scripts = [
        "_layouts/15/init.debug.js",
        "_layouts/15/MicrosoftAjax.js",
        "_layouts/15/sp.core.debug.js",
        "_layouts/15/sp.runtime.debug.js",
        "_layouts/15/sp.debug.js"
    ];

    return new Promise((resolve, reject) => {
        scripts.reduce((result, currentValue, currentIndex, array) => {
            return result.then((loadedScript) => {
                if(loadedScript) vm.runInThisContext(loadedScript);
                return new Promise((res, rej) => {
                    const lib = require("http");
                    const request = lib.get(url.resolve(serverAbsoluteUrl,currentValue), (response) => {
                        const body = [];
                        response.on("data", (chunk) => body.push(chunk));
                        response.on("end", () => res(body.join("")));
                    });
                    request.on("error", (err) => reject(err));
                });
            });
        }, Promise.resolve()).then((loadedScript) => {
            if(loadedScript) vm.runInThisContext(loadedScript);
            resolve(true);
        });
    }); 
}
