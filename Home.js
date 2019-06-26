(function () {
    "use strict";

    var messageBanner;
    var authenticator;
    var googleToken;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the FabricUI notification mechanism and hide it
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();

            // Determine if we are running inside of an authentication dialog
            // If so then just terminate the running function
            if (OfficeHelpers.Authenticator.isAuthDialog()) {
                // Adding code here isn"t guaranteed to run as we need to close the dialog
                // Currently we have no realistic way of determining when the dialog is completely
                // closed.
                return;
            }

            // Create a new instance of Authenticator
            authenticator = new OfficeHelpers.Authenticator();

            // Register our providers accordingly
            authenticator.endpoints.registerGoogleAuth('478306342633-o9go66u2bf65atn2lgmcfjlcbo9h66ag.apps.googleusercontent.com');
            
            $('#button-text').text("Authenticate!");
            $('#button-desc').text("Authenticate via Google account.");
                
            // Add a click event handler for the highlight button.
            $('#highlight-button').click(openGoogle);
        });
    };

    function loadSampleData() {
        var values = [
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)]
        ];

        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {
            // Create a proxy object for the active sheet
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            // Queue a command to write the sample data to the worksheet
            sheet.getRange("B3:D5").values = values;

            // Run the queued-up commands, and return a promise to indicate task completion
            return ctx.sync();
        })
        .catch(errorHandler);
    }

    function openGoogle() {
        var initialUrl = window.location.origin + '/test1/Functions/AuthDialog.html#initialize=true';

        Office.context.ui.displayDialogAsync(initialUrl, { width: 30, height: 50 }, function (result) {
            if (result.status === Office.AsyncResultStatus.Failed) {
            }
            else {
                let dialog = result.value;
                dialog.addEventHandler(Office.EventType.DialogMessageReceived, function(args) {
                    var resultStr = args.message;
                    var result = JSON.parse(resultStr);
                    if (result.accessToken) {
                        googleToken = result;
                        dialog.close();
                    }
                });
                dialog.addEventHandler(Office.EventType.DialogEventReceived, function (args) {
                    var result = args.message;
                });
            }
        });
    }

    function authenticate() {
        authenticator
            .authenticate(OfficeHelpers.DefaultEndpoints.Google, true)
            .then(function (token, q, w, e, r, t, y, u) {
                console.log('_GOOGLE_TOKEN: ', token);
            })
            .catch(function (error) {
                console.log('_GOOGLE_TOKEN: ', error);
            });
    }

    function displaySelectedCells() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    showNotification('Error', result.error.message);
                }
            });
    }

    // Helper function for treating errors
    function errorHandler(error) {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        showNotification("Error", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
