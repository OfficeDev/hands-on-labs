$(function () {
    $('#AVChatVideo').click(startConversation);
    var SkypeWebApp;
    var SkypeApi;
    //var sip = getUrlVars()["sip"];
    //var audioOnly = getUrlVars()["audioOnly"];
    var muted = false;
    //var video = !(audioOnly == "true");
    Skype.initialize({
        apiKey: config.apiKey
    }, function (api) {
        SkypeApi = api;
        SkypeWebApp = new api.application({
            settings: {
                supportsText: true,
                supportsHtml: false,
                supportsMessaging: true,
                supportsAudio: true,
                supportsVideo: true,
                supportsSharing: false
            }
        });
        console.log(SkypeWebApp);
        // whenever client.state changes, display its value
        SkypeWebApp.signInManager.state.changed(function (state) {
            console.log("Skype Client state changed to: " + state);
        });
        var options = {
            "auth": null,
            "client_id": "98267106-694b-4df2-8e06-fbbafd8a90e7",
            "origins": ["https://webdir.online.lync.com/autodiscover/autodiscoverservice.svc/root"],
            "cors": true,
            "version": 'secondonlineapp/1.0.0.0',
            "redirect_uri": "/Home.html",
        };
        if (SkypeWebApp.signInManager.state() == 'SignedOut') {
            SkypeWebApp.signInManager.signIn(options).then(function () {
                console.log('signed in');
                avListener();
            },
            function (error) {
                console.log('sign-in' + error);
            });
        }
    });
    function avListener() {
        SkypeWebApp.conversationsManager.conversations.added(function (conversation) {
            conversation.selfParticipant.audio.state.changed(function (newState, reason, oldState) {
                if (newState == 'Notified') {
                    console.log("Audio notified");
                }
                else if (newState == 'Connected') {
                    console.log("Connected to Audio service");
                }
                else if (newState == "Disconnected") {
                    console.log("Disconnected from audio service");
                }
            });
            conversation.selfParticipant.video.state.changed(function (newState, reason, oldState) {
                if (newState == 'Notified') {
                    console.log("Video notified");
                }
                else if (newState == 'Connected') {
                    console.log("Connected to Video service");
                }
                else if (newState == "Disconnected") {
                    console.log("Disconnected from video service");
                }
            });

            function onAudioNotified() {
                //var name = conversation.participants(0).person.displayName();
                if (conversation.state() == 'Connected') {
                    conversation.audioService.accept();
                }
                else {
                    if (confirm('Accept an audio call from ' + name + '?')) {
                        console.log('accepting the audio call');
                        conversation.audioService.accept();
                    }
                    else {
                        console.log('declining the incoming audio request');
                        conversation.audioService.reject();
                    }
                }
            }

            function onAudioVideoNotified() {
                if (conversation.selfParticipant.video.state() == "Notified") {
                    onVideoNotified();
                }
                else {
                    onAudioNotified();
                }
            }
        });
    }
    function findPerson(uri) {
        var searchQuery = SkypeWebApp.personsAndGroupsManager.createPersonSearchQuery();
        searchQuery.text(uri);
        return searchQuery.getMore().then(function (results) {
            return results[0].result;
        });
    }
    function addParticipant(conv, uri) {
        return findPerson(uri).then(function (person) {
            conv.participants.add(conv.createParticipant(person));
        });
    }
    //function startConversation() {
    //    var p = findPerson('troberts@claritycon.com').then(function (person) {;

    //        console.log('Creating conversation');
    //        var conversation = SkypeWebApp.conversationsManager.createConversation();
    //        console.log('Creating participant (' + sip + ')');
    //        var convParticipant = conversation.createParticipant(person)
    //        console.log('Adding participant to conversation');
    //        conversation.participants.add(convParticipant);
    //        console.log('Adding conversation to manager');
    //        SkypeWebApp.conversationsManager.conversations.add(conversation);
    //    });
    //}

    function startConversation() {
        var sip = 'sip:barmstrong@danewman.onmicrosoft.com';
        var person;
        //console.log('looking up based on sip');
        GetContactFromName(sip).then(function (results) {
            results.forEach(function (result) {
                person = result.result;
            });

            //console.log('person created: ' + person);
            var conversation = SkypeWebApp.conversationsManager.createConversation();
            //console.log('conversation created');
            var convParticipant = conversation.createParticipant(person)
            //console.log('participant created: ' + convParticipant);
            conversation.participants.add(convParticipant);
            //console.log('participant added');
            SkypeWebApp.conversationsManager.conversations.add(conversation);
            conversation.chatService.start();
            conversation.audioService.start();
            //console.log('conversation at this point: ' + conversation);
        });
    }

    function GetContactFromName(contactSIP) {
        var query = SkypeWebApp.personsAndGroupsManager.createPersonSearchQuery();
        query.text(contactSIP);
        query.limit(1);
        //console.log('returning search results');
        return query.getMore();
    }
});