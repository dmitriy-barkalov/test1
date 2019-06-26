// The initialize function must be run each time a new page is loaded.
(function () {
    "use strict";

    Office.initialize = function (reason) {
        $(document).ready(function () {
            var params = parseParameters(window.location.hash.substring(1));

            if (params.initialize) {
                var initialUrl = 'https://accounts.google.com/o/oauth2/v2/auth?' +
                    'scope=https%3A%2F%2Fwww.googleapis.com%2Fauth%2Fdrive.metadata.readonly&' +
                    'include_granted_scopes=true&' +
                    'redirect_uri=' + window.location.origin + window.location.pathname + '&' +
                    'response_type=token&' +
                    'client_id=478306342633-o9go66u2bf65atn2lgmcfjlcbo9h66ag.apps.googleusercontent.com';

                window.location.href = initialUrl;
            } else if (params.access_token) {
                var result = {
                    result : 'success',
                    accessToken: params.access_token,
                    expiry : Date.now() + params.expires_in,
                    tokenType : params.token_type
                }
                Office.context.ui.messageParent(JSON.stringify(result));
            } else {
                Office.context.ui.messageParent(JSON.stringify({
                    result : 'error',
                    description: 'Something Wrong',
                    href: window.location.href
                }));
            }
        });
    };

    function parseParameters(params) {
        var vars = params.split('&');

        var authParams = {};
        for (var i = 0; i < vars.length; i++) {
            var pair = vars[i].split('=');
            authParams[decodeURIComponent(pair[0])] = decodeURIComponent(pair[1]);
        }

        return authParams;
    }
})();