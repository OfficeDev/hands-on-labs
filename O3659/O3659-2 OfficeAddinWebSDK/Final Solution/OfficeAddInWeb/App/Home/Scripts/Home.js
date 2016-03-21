/// <reference path="../App.js" />
var access_token;
var registeredListeners = registeredListeners || [];

registeredListeners.forEach(function (listener) {
    listener.dispose();
});

registeredListeners = [];

(function () {
    "use strict";
    var client;
    var apiManager;
    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {           
            $('#signinbtn-email').click(showSkypeLogin);
            $('#selectedChat').click(startConversation);
            $('#closeChat').click(endConversation);
            pageLoading();
            Office.context.document.addHandlerAsync("documentSelectionChanged", excelClickHandler, function (result) {
            });

            if (/^#access_token=/.test(location.hash)) {               
                access_token = location.hash.match(/\#(?:access_token)\=([\S\s]*?)\&/)[1];
                console.log('accessToken: ' + access_token);
                $('#accessToken').val(access_token);
            }
            else {
                $('#main-content').show();
                $('#content-loading').hide();
}
            Skype.initialize({
                apiKey: config.apiKeyCC
            }, function (api) {
                apiManager = api;
                client = apiManager.UIApplicationInstance;
                client.signInManager.state.changed(function (state) {
                    $('#loginState').text(state);
                });
                if (access_token) {
                    signIn();
                }

               
            });
        });
    };
    
    function conversationHandler() {
        var test = '';
        client.conversationsManager.conversations.added(function (conversation) {
            conversation.selfParticipant.chat.state.when('Notified', function () {
                $('#conversation').show();
                $('.contact-info').hide();
                var id = conversation.participants(0).person.id();
                var container = document.getElementById(id);
                if (!container) {
                    container = document.createElement('div');
                    container.id = id;
                    document.querySelector('#conversations').appendChild(container);
                }
                else {
                    document.querySelector('#conversations').removeChild(container);
                }
                var promise = apiManager.renderConversation(container, {
                    modalities: ['Chat'],
                    participants: [id]
                });
                monitor('start conversation', promise);
            });
        });

    }

    function startConversation() {
        $('#conversation').show();
        $('.contact-info').hide();
        var chatSip = $('#selectedSIP').val();
        var uris = [chatSip];
        var container = document.getElementById(chatSip);
        if (!container) {
            container = document.createElement('div');
            container.id = chatSip;
            document.querySelector('#conversations').appendChild(container);
        }
        else {
            document.querySelector('#conversations').removeChild(container);
        }
            var promise = apiManager.renderConversation(container, { modalities: ['Chat'], participants: uris });
            monitor('start conversation', promise);
    }

    function endConversation() {
        $('#conversation').hide();
        $('.contact-info').show();
        apiManager.UIApplicationInstance.conversationsManager.conversations.get().then(function (conversationsArray) {
            if (conversationsArray && conversationsArray.length > 0) {
                conversationsArray.forEach(function (element, index, array) {
                    console.log("Closing existed conversation...");                   
                    var convo = apiManager.UIApplicationInstance.conversationsManager.conversations(0);
                    convo.leave();
                    apiManager.UIApplicationInstance.conversationsManager.conversations.remove(element);
                });
            }
        });
    }

    function monitor(title, promise) {
        console.log(title, 'started');
        promise.then(function (res) {
            console.log(title, 'succeeded', res);
        }, function (err) {
            console.log(title, 'failed', err && err.stack || err);
            alert(title + ' failed:' + err);
        });
    }

    function showContacts() {
        $('.signIn').hide();
        $('.loggedIn').show();
        conversationHandler();
        $('#main-content').show();
        $('#content-loading').hide();
    }

    function signIn() {
        console.log('starting sign-in');
       var params =
    {
    "auth": null,
    "client_id": "98267106-694b-4df2-8e06-fbbafd8a90e7",
    "origins": ["https://webdir.online.lync.com/autodiscover/autodiscoverservice.svc/root"],
    "cors": true,
    "version": 'secondlineapp/1.0.0.0'
    };

        console.log(params);
        console.log('user state: ' + client.signInManager.state());
        if (client.signInManager.state() == 'SignedOut') {
            client.signInManager.signIn(
                params
            ).then(function () {
                console.log('logged in');
                showContacts();
            }, function (error) {
                console.log(
                    'Error signing in: '+ error || 'Cannot sign in');
                
            });
        }
        else {
            console.log('already logged in');
            showContacts();
        }
    }

    function displayError(error) {
        $('#error').text(error || 'error');
    }

    function showSkypeLogin() {
        var url = config.authLink +
               '&client_id=' + config.clientId +
               '&redirect_uri=' + config.redirect_uri +
               '&resource=' + config.authResource;
        console.log('login url: ' + url);
        location.assign(url);
    }

    function getMyEmployeeInfo() {
        var me = client.personsAndGroupsManager.mePerson;
        $('#empName').text(me.displayName());
        $('#empEmail').text(me.email());
        me.avatarUrl.changed(function (url) {
            var aUrl = url;
            //TODO: avatar URL is currently a 401 UNAUTHORIZED
            //$('#empImg').attr('src', url);
            $('#empImg').attr('src', 'Images/default.png');
        });
        me.avatarUrl.subscribe();
    }

    function subscribeToPresenceChanges() {
        var me = client.personsAndGroupsManager.mePerson;
        me.status.changed(function (newStatus) {
            $('#presenceState').text(newStatus);
            $('#presenceIcon').attr('src', setPresenceIcon(newStatus));
        });
        me.activity.changed(function (newActivity) { });
    }

    function getSelectedEmployeeInfo(email) {
        if (email) {
            var emp = email;
            var query = client.personsAndGroupsManager.createPersonSearchQuery();
            query.text(emp);
            query.limit(1);
            query.getMore().then(function (results) {
                $('#contactInfo').show();
                results.forEach(function (item) {
                    var person = item.result;
                    $('#selectedName').text(person.displayName());
                    var avatar = person.avatarUrl();
                    //TODO: avatar URL is currently a 401 UNAUTHORIZED
                    //$('#selectedPhoto').attr("src", avatar);
                    $('#selectedPhoto').attr('src', 'Images/default.png');
                    var sip = person.id();
                    $('#selectedSIP').val(sip);
                    person.emails.get().then(function (e) {
                        e.forEach(function (email) {
                            $('#selectedEmail').text(email.emailAddress());
                        });
                    });
                    person.status.get().then(function (s) {
                        $('#selectedStatus').text(s);
                        $('#selectedPresenceIcon').attr('src', setPresenceIcon(s));
                    });
                    person.status.changed(function (newStatus) {
                        $('#selectedStatus').text(newStatus);
                        $('#selectedPresenceIcon').attr('src', setPresenceIcon(newStatus));
                    });
                    person.status.subscribe();
                    $('#selectedCall').attr('href', 'Callto:' + sip);
                    $('#selectedVideo').attr('href', 'Callto:' +sip);
                });
            });
        }
    }
    function pageLoading() {
        $('#main-content').hide();
        $('#content-loading').show();
    }
    function excelClickHandler(eventArgs) {
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

    function setPresenceIcon(state) {
        switch (state) {
            case "Online":
                return "Images/available.png";
            case "Offline":
            case "Unknown":
            case "Hidden":
            default:
                return "Images/unknown.png";
            case "Busy":
                return "Images/busy.png";
            case "Idle":
            case "IdleOnline":
            case "Away":
            case "BeRightBack":
                return "Images/away.png";
            case "DoNotDisturb":
                return "Images/do-not-disturb.png";

        }
    }

})();