$(function () {
    var SkypeWebApp, SkypeApi;
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
                supportsSharing: false,
            }
        });

        SkypeWebApp.signInManager.state.changed(function (state) {
            console.log("Skype Client state changed to: " + state);
        });

        var options =
        {
            "auth": null,
            "client_id": "98267106-694b-4df2-8e06-fbbafd8a90e7",
            "origins": ["https://webdir.online.lync.com/autodiscover/autodiscoverservice.svc/root"],
            "cors": true,
            "version": 'secondonlineapp/1.0.0.0',
            "redirect_uri": "/9c967f6b-a846-4df2-b43d-5167e47d81e1/oauth2/token/index.html",
        };
        SkypeWebApp.signInManager.signIn(options).then(function () {
            console.log('Signed In');
            //subscribeToChatEvents();
            startConversation();
        },
        function (error) {
            console.log('sign-in' + error);
        });

    }, function (err) {
        console.log(err);
        alert('Cannot load the SDK.');
    });

    function findPerson() {
        //var sip = "sip:marysmith@danewman.onmicrosoft.com";
        //var searchQuery = SkypeWebApp.personsAndGroupsManager.createPersonSearchQuery();
        //searchQuery.text(sip);
        //return searchQuery.getMore().then(function (results) {
        //    console.log('results');
        //    return results[0].result;
        //});
    }
    function startConversation() {

        GetContactFromName('sip:troberts@claritycon.com').then(function (results) {
            results.forEach(function (result) {
                person = result.result;
            });
            console.log('creating conversation');
            var conversation = SkypeWebApp.conversationsManager.createConversation();
            var convParticipant = conversation.createParticipant(person)
            subscribeToParticipant(convParticipant);
            conversation.participants.add(convParticipant);
            SkypeWebApp.conversationsManager.conversations.add(conversation);
            console.log('created participant, added them, added convo to manager');
            console.log('wait_1 (5 seconds) beginning');
            setTimeout(function () {
                console.log('wait_1 finished');
                conversation.videoService.start();
                console.log(conversation);
                console.log('wait_2 (5 seconds) beginning');
                setTimeout(function () {
                    console.log('wait_2 finished');
                    conversation.selfParticipant.video.state.changed(function (newState) {
                        if (newState == 'Connected') {
                            conversation.selfParticipant.video.channels(0).stream.source.sink.container.set(document.getElementById("render-self-window"));
                            conversation.selfParticipant.video.channels(0).isStarted.set(true);
                            convParticipant.video.state.changed(function (state) {
                                if (state == 'Connected') {
                                    console.log("The remote participant video has connected");
                                    convParticipant.video.channels(0).stream.source.sink.container.set(document.getElementById("render-participant-window"));
                                    convParticipant.video.channels(0).isStarted.set(true);
                                    console.log(conversation);
                                    console.log("The remote participant sink set has completed");
                                }
                            });
                        }
                    });
                }, 5000);
            }, 5000);
        });
        //var person;
        //console.log('starting convo');
        //findPerson().then(function (result) {
        //    person = result;
        //    console.log('person: ' + person);
        //    var conversation = SkypeWebApp.conversationsManager.createConversation();
        //    var par = conversation.createParticipant(person);
        //    conversation.participants.add(par);
        //    SkypeWebApp.conversationsManager.conversations.add(conversation);
        //    conversation.chatService.start();
        //    conversation.chatService.state.changed(function (newState) {
        //        if (newState == "Connected") {
        //            conversation.audioService.start();
        //        }
        //    });
        //    conversation.selfParticipant.audio.state.changed(function (newState) {
        //        console.log(newState);
        //        if (newState == 'Notified') {
        //            if (confirm('accept audio?')) {
        //                conversation.audioService.accept();
        //                console.log('audioservice accepted');
        //            }
        //            else {
        //                conversation.audioService.reject();
        //            }
        //        }
        //    });
        //});
    }

    function GetContactFromName(contactSIP) {
        var query = SkypeWebApp.personsAndGroupsManager.createPersonSearchQuery();
        query.text(contactSIP);
        query.limit(1);
        return query.getMore();
    }

    //function subscribeToChatEvents() {
    //    var timer;
    //    SkypeWebApp.conversationsManager.conversations.added(function (conversation) {
    //        //conversation.selfParticipant.audio.state.changed(function (newState, reason, oldState) {
    //        //    if (newState == 'Notified') {
    //        //        console.log("Audio notified");
    //        //        conversation.audioService.accept();
    //        //    }
    //        //    else if (newState == 'Connected') {
    //        //        console.log("Connected to Audio service");
    //        //        renderAudioService(conversation);
    //        //    }
    //        //    else if (newState == "Disconnected") {
    //        //        console.log("Disconnected from audio service");
    //        //    }
    //        //});
    //        //conversation.selfParticipant.video.state.changed(function (newState, reason, oldState) {
    //        //    if (newState == 'Notified') {
    //        //        conversation.videoService.accept();
    //        //        console.log("Video notified");
    //        //    }
    //        //    else if (newState == 'Connected') {
    //        //        console.log("Connected to Video service");
    //        //        renderVideoService(conversation);
    //        //    }
    //        //    else if (newState == "Disconnected") {
    //        //        console.log("Disconnected from video service");
    //        //    }
    //        //});
    //        console.log(person);
    //        //conversation.participants(0).video.state.changed(function (newState) {
    //        //    console.log('does this ever work?');
    //        //});
    //    });
    //////}

    function subscribeToParticipant(participant) {
        var timer;
        participant.video.state.changed(function (newState, reason, oldState) {
            if (oldState == 'Connecting' && newState == 'Disconnected') {
                timer = setInterval(function () { console.log('test'); }, 1000);
            }

            if (newState == 'Connected') {
                clearInterval(timer);
            }

            if (newState == 'Disconnected') {
                console.log('Old state: ' + oldState + ' New state: ' + newState + ' Reason: ' + reason);
            }
        });
    }

});