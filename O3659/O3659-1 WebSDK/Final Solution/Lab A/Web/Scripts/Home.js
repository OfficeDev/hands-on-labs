var client;
var apiManager;
var access_token;
var contactNo = 0;

$(function () {
    Skype.initialize({
        apiKey: config.apiKeyCC
    }, function (api) {
        apiManager = api;
        client = apiManager.UIApplicationInstance;
        client.signInManager.state.changed(function (state) {
            console.log('logged in user: '+ state);
        });
        signIn();
    });

    function toggleChat(sip) {
        startConversation(sip);
    }

    function toggleContacts() {
        endConversation();
    }

    function displayContactCard(contactCardID) {
        //get user data from the Skype logins here
        var cardName = $('#' + contactCardID + ' .contactName').html();
        var cardPresence = $('#' + contactCardID + ' .contactPresence').html();
        var cardSIP = $('#' + contactCardID + ' .contactSIP').val();
        $('#ContactCard #ContactPresence').html(cardPresence);
        $('#ContactCard .contactCardSIP').val(cardSIP);
        $('#ContactCard').dialog({
            dialogClass: "no-close",
            title: cardName,
            draggable: false,
            resizable: false,
            position: { my: "right", at: "left", of: $('#' + contactCardID) },
            buttons: [
              {
                  icons: {
                      primary: "ui-icon-chat"
                  },
                  click: function () {
                      $(this).dialog("close");
                      //replace the contact list with the chat window. Add the call and video buttons there as well
                      $('#ChatContact').val(cardName); //we'd store actual info here, not just the name
                      toggleChat(cardSIP);
                  }
              },
              {
                  icons: {
                      primary: "ui-icon-call"
                  },
                  click: function () {
                      $(this).dialog("close");
                      var params = "sip=" + cardSIP + "&audioOnly=true";
                      launchAVWindow(params);
                      //pop open the call window for audio. You can end the call or initiate video from here
                  },
              },
              {
                  icons: {
                      primary: "ui-icon-video"
                  },
                  click: function () {
                      $(this).dialog("close");
                      var params = "sip=" + cardSIP + "&audioOnly=false";
                      launchAVWindow(params);
                      //pop open the call window for full audio and video. You can end the call or turn off video, etc.
                  }
              }
            ]
        });
    }

    function launchAVWindow(params) {
        location.assign('/AVDemo.html?' + params);
    }

    function getSelectedEmployeeInfo(email) {
        if (email) {
            console.log('employee info: ' + email);
            var emp = email;
            var query = client.personsAndGroupsManager.createPersonSearchQuery();
            query.text(emp);
            query.limit(1);
            query.getMore().then(function (results) {
                results.forEach(function (item) {
                    var person = item.result;
                    var avatar = person.avatarUrl();
                    var sip = person.id();
                    console.log('sip:' + sip);
                    $('#ContactList').append('<div id="Contact' + contactNo + '" class="contact" data-sip="' + sip + '">' +
                        '<input type="hidden" value="' + sip + '" class="contactSIP"/>' +
                        '<div class="contactName">' + person.displayName() + '</div>' +
                        '<div><img class="contactListAvatar" src="Images/default.png" /></div>' +
                        '<div class="contactPresence"></div>' +
                        '</div>');
                    contactNo++;
                    person.status.get().then(function (s) {
                        console.log('initial status:' + s);
                        $('[data-sip="' + person.id() + '"] .contactPresence').html(s);
                    });
                    person.status.changed(function (newStatus) {
                        console.log('new status:' + newStatus);
                        $('[data-sip="' + person.id() + '"] .contactPresence').html(newStatus);
                        $('#ContactsLoadingGif').hide();
                        $('#ContactList').show();
                    });
                    $('#selectedCall').attr('href', 'Callto:' + sip);
                    $('#selectedVideo').attr('href', 'Callto:' + sip);
                    person.status.subscribe();
                });
            });
        }
    }

    function conversationHandler() {
        var test = '';
        client.conversationsManager.conversations.added(function (conversation) {
            conversation.selfParticipant.chat.state.when('Notified', function () {
                console.log('new conversation notification');
                $('#ContactList').hide();
                $('#ChatWindow').show();
                $('#ChatControlsArea').show();
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

    function startConversation(sip) {
        $('#ContactList').hide();
        $('#ChatWindow').show();
        $('#ChatControlsArea').show();
        console.log(sip);
        var chatSip = sip;
        var uris = [chatSip];
        var container = document.getElementById(chatSip);
        if (!container) {
            container = document.createElement('div');
            container.id = chatSip;
            document.querySelector('#conversations').appendChild(container);
        }
        var promise = apiManager.renderConversation(container, { modalities: ['Chat'], participants: uris });
        monitor('start conversation', promise);
    }

    function endConversation() {
        apiManager.UIApplicationInstance.conversationsManager.conversations.get().then(function (conversationsArray) {
            if (conversationsArray && conversationsArray.length > 0) {
                conversationsArray.forEach(function (element, index, array) {
                    console.log("Closing existing conversation...");
                    var convo = apiManager.UIApplicationInstance.conversationsManager.conversations(0);
                    convo.leave();
                    apiManager.UIApplicationInstance.conversationsManager.conversations.remove(element);
                    $('#conversations').empty();
                });
            }
        });
        $('#ChatWindow').hide();
        $('#ChatControlsArea').hide();
        $('#ContactList').show();
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



    function signIn() {
        var params =
     {
         "auth": null,
         "client_id": config.clientId,
         "origins": ["https://webdir.online.lync.com/autodiscover/autodiscoverservice.svc/root"],
         "cors": true,
         "version": 'secondonlineapp/1.0.0.0',
         "redirect_uri": "/9c967f6b-a846-4df2-b43d-5167e47d81e1/oauth2/token/index.html",
     };
        console.log(params);
        if (client.signInManager.state() == 'SignedOut') {
            client.signInManager.signIn(
                params
            ).then(function () {
                console.log('logged in');
                loadContacts();
                showLoggedInUserData();
                conversationHandler();
                client.personsAndGroupsManager.mePerson.status.changed(function (newStatus) {
                    console.log('logged in status: ' + newStatus);                
                });
            }, function (error) {
                console.log(
                    'Error signing in: ' + error || 'Cannot sign in');
            });
        }
        else {
            console.log('already logged in');
            loadContacts();
            showLoggedInUserData();
            conversationHandler();
        }
    }

    function loadContacts() {
        var contactList = client.personsAndGroupsManager.all.persons.get().then(function (persons) {
            persons.forEach(function (person) {
                person.id.get().then(function (id) {
                    getSelectedEmployeeInfo(id.replace('sip:', ''));
                });
            });
        });
    }

    function showLoggedInUserData() {
        $('#LoadingGif').hide();
        $('#SkypeContent').show();
        $('#Content').show();
    }

    function getMyEmployeeInfo() {
        var me = client.personsAndGroupsManager.mePerson;
        me.displayName.get().then(function (value) {
            consol.log(value);
            $('#UserName').text(value);
        });
        console.log("displayName: " + me.displayName());
        me.email.get().then(function (value) {
            consol.log(value);
            $('#UserEmail').text(value);
        });
        console.log("email: " + me.email());
        me.avatarUrl.changed(function (url) {
            var aUrl = url;
            //TODO: avatar URL is currently a 401 UNAUTHORIZED
            //$('#empImg').attr('src', url);
            $('#UserAvatar').attr('src', 'Images/default.png');
        });
        me.avatarUrl.subscribe();
    }
    function subscribeToPresenceChanges() {
        var me = client.personsAndGroupsManager.mePerson;
        me.status.changed(function (newStatus) {
            $('#UserPresenceIcon').attr('src', setPresenceIcon(newStatus));
        });
        me.activity.changed(function (newActivity) { });
    }

    function setPresenceIcon(state) {
        switch (state) {
            case "Online":
                return "Images/available.png";
            case "Busy":
                return "Images/busy.png";
            case "Idle":
            case "IdleOnline":
            case "Away":
            case "BeRightBack":
                return "Images/away.png";
            case "DoNotDisturb":
                return "Images/do-not-disturb.png";
            case "Offline":
            case "Unknown":
            case "Hidden":
            default:
                return "Images/unknown.png";
        }
    }

    //listener on contact cards
    $('#CCContainer').on('click', '.contact', function () {
        displayContactCard($(this).attr('id'));
    });

    //listener on chat window controls
    //audio
    $('#CCContainer').on('click', '#ChatCall', function () {
        var params = "sip=" + $('#conversations > div').attr('id') + "&audioOnly=true";
        launchAVWindow(params);
    });
    //video
    $('#CCContainer').on('click', '#ChatVideo', function () {
        var params = "sip=" + $('#conversations > div').attr('id') + "&audioOnly=false";
        launchAVWindow(params);
    });
    //end chat
    $('#CCContainer').on('click', '#ChatClose', function () {
        $('#ChatContact').val(''); //clear the current conversation data
        toggleContacts();        
    });

});