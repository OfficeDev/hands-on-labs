/// <reference path="../App.js" />
var access_token;
(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            $('#signinbtn-email').click(showSkypeLogin);
            Office.context.document.addHandlerAsync("documentSelectionChanged", myHandler, function (result) {
            });

            if (/^#access_token=/.test(location.hash)) {
                access_token = location.hash.match(/\#(?:access_token)\=([\S\s]*?)\&/)[1];
                $('#accessToken').val(access_token);
            }
            Skype.initialize({
                apiKey: config.apiKeyCC,
            }, function (api) {
                Office.context.skypeWebApp = api.UIApplicationInstance;
                Office.context.skypeApi = api;
                Office.context.skypeWebApp.signInManager.state.changed(function (state) {
                    $('#loginState').text(state);
                });
                if (access_token) {
                    pageLoading();
                    signIn();
                }
            });
        });
    };



    function signIn() {
        var params =
     {
         "auth": access_token,
         "client_id": config.clientId,
         "origins": ["https://webdir.tip.lync.com/autodiscover/autodiscoverservice.svc/root"],
         "cors": true,
         "version": "sdk-samples/1.0.0"
     };
        if (Office.context.skypeWebApp.signInManager.state() == 'SignedOut') {
            Office.context.skypeWebApp.signInManager.signIn(
                params
            ).then(function () {
                location.assign("Home.html");
            }, function (error) {
                console.log(error || 'Cannot sign in');

            });
        }
        else {
        }
    }

    function displayError(error) {
        $('#error').text(error || 'error');
    }

    function showSkypeLogin() {
        location.assign('https://login.windows-ppe.net/common/oauth2/authorize?response_type=token' +
               '&client_id=a134ad4c-f9a4-48dc-be60-9de6b3124166' +
               '&redirect_uri=https://localhost:44300/App/Home/Login.html?_host_Info=Excel|Win32|16.01|en-US' +
               '&resource=https://webdir0d.tip.lync.com');
    }

    function subscribeToPresenceChanges() {
        var me = Office.context.skypeWebApp.personsAndGroupsManager.mePerson;
        me.status.changed(function (newStatus) {
            $('#presenceState').text(newStatus);
            $('#presenceIcon').attr('src', setPresenceIcon(newStatus));
        });
        me.activity.changed(function (newActivity) { });
    }

    function pageLoading() {
        $('#content-main').hide();
        $('#content-loading').show();
    }

    function myHandler(eventArgs) {
        eventArgs.document.getSelectedDataAsync(Office.CoercionType.Text, function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                $('#selectedText').text('broke');
            }
            else {
                $('#selectedText').text(asyncResult.value);
                getSelectedEmployeeInfo(asyncResult.value);
            }
        });
    }



})();