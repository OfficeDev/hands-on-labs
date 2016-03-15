function getUrlVars() {
    var vars = [], hash;
    var hashes = window.location.href.slice(window.location.href.indexOf('?') + 1).split('&');
    for (var i = 0; i < hashes.length; i++) {
        hash = hashes[i].split('=');
        vars.push(hash[0]);
        vars[hash[0]] = hash[1];
    }
    return vars;
}

$(function () {
    var SkypeWebApp;
    var SkypeApi;
    var sip = getUrlVars()["sip"];
    var audioOnly = getUrlVars()["audioOnly"];
    var muted = false;
    var video = false;
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
        console.log(options);
        SkypeWebApp.signInManager.signIn(options).then(function () {
            console.log('signed in');
            //subscribeToChatEvents();
            startConversation(sip);
        },
        function (error) {
            console.log('sign-in' + error);
        });

    }, function (err) {
        console.log(err);
        alert('Cannot load the SDK.');
    });

    if (audioOnly == "true") {
        $('#AVChatVideo').css('background', 'url(../Images/video.png) no-repeat');
        $('#AVChatVideo').css('background-size', '50px 50px');
    } else {
        //start video
        $('#AVChatVideo').css('background', 'url(../Images/video-off.png) no-repeat');
        $('#AVChatVideo').css('background-size', '50px 50px');
        $('#AVWindowSelfView').css('background-image', 'url(../Images/self-video.png)');
    }

    function startConversation(sip) {
        var person;
        console.log('looking up based on sip');
        GetContactFromName(sip).then(function (results) {
            results.forEach(function (result) {
                person = result.result;
            });

            console.log('person created: ' + person);
            var conversation = SkypeWebApp.conversationsManager.createConversation();
            console.log('conversation created');
            var convParticipant = conversation.createParticipant(person)
            console.log('participant created: ' + convParticipant);
            conversation.participants.add(convParticipant);
            console.log('participant added');
            SkypeWebApp.conversationsManager.conversations.add(conversation);
            console.log('converation added to conversations manager');
            console.log('conversation at this point: ' + conversation);
            conversation.videoService.start();
            conversation.selfParticipant.video.state.changed(function (newState) {
                if (newState == 'Connected') {
                    conversation.selfParticipant.video.channels(0).stream.source.sink.container.set(document.getElementById("render-self-window"));
                    conversation.selfParticipant.video.channels(0).isStarted.set(true);
                    convParticipant.video.state.changed(function (state) {
                    if (state == 'Connected') {
                            console.log("The remote participant video has connected");
                            convParticipant.video.channels(0).stream.source.sink.container.set(document.getElementById("render-participant-window"));
                            convParticipant.video.channels(0).isStarted.set(true);
                            console.log("The remote participant sink set has completed");
                        }
                    });
                }
            });
        });
    }

    function GetContactFromName(contactSIP) {
        var query = SkypeWebApp.personsAndGroupsManager.createPersonSearchQuery();
        query.text(contactSIP);
        query.limit(1);
        console.log('returning search results');
        return query.getMore();
    }

    //wire to click event
    function startAudioChat(sip) {
        var conversation = startConversation(sip);
    }

    //wire to click event
    function startVideoChat(sip) {
        var conversation = startConversation(sip);
        conversation.videoService.start();
    }

    //listener event should be called once skype is initialized
    function subscribeToChatEvents() {
        SkypeWebApp.conversationsManager.conversations.added(function (conversation) {
            conversation.selfParticipant.audio.state.changed(function (newState, reason, oldState) {
                if (newState == 'Notified') {
                    console.log("Audio notified");
                    conversation.audioService.accept();
                }
                else if (newState == 'Connected') {
                    console.log("Connected to Audio service");
                    renderAudioService(conversation);
                }
                else if (newState == "Disconnected") {
                    console.log("Disconnected from audio service");
                }
            });
            conversation.selfParticipant.video.state.changed(function (newState, reason, oldState) {
                if (newState == 'Notified') {
                    conversation.videoService.accept();
                    console.log("Video notified");
                }
                else if (newState == 'Connected') {
                    console.log("Connected to Video service");
                    renderVideoService(conversation);
                }
                else if (newState == "Disconnected") {
                    console.log("Disconnected from video service");
                }
            });
        });
    }

    //method to be called once video has been notified as "connected"
    function renderVideoService(conversation) {
        var container = '#AVContent';

        //open self camera
        var selfChannel = conversation.selfParticipant.video.channels(0);
        selfChannel.stream.source.sink.container.set($("#AVWindowSelfView"));

        //open participant camera stream
        console.log(conversation.participants());
        var participant = conversation.participants(0);
        var participantChannel = participant.video.channels(0);
        participantChannel.stream.source.sink.container.set($("#AVWindowContactView"));

        conversation.audioService.start();
        conversation.chatService.start();
    }

    //method to be called once audio has been notified as "connected"
    function renderAudioService(conversation) {
        conversation.audioService.start();
    }

    //end conversation, wire up to button event
    function endConversation() {
        //stop all services
        SkypeWebApp.conversationsManager.conversations.added(function (conversation) {
            conversation.videoService.stop();
            conversation.audioService.stop();
            conversation.chatService.stop();
            SkypeWebApp.conversationsManager.conversations.remove(conversation);
        });
    }

    $('#AVChatVideo').click(function () {
        if (!video) {
            //this is where we would escalate to video
            $('#AVChatVideo').css('background', 'url(../Images/video-off.png) no-repeat');
            $('#AVChatVideo').css('background-size', '50px 50px');
            $('#AVWindowSelfView').css('background-image', 'url(../Images/self-video.png)');
            //$('#AVWindowContactView').css('background-image', 'url(../Images/participant-video.png)');
            video = true;
        } else {
            //turn off self-video
            $('#AVChatVideo').css('background', 'url(../Images/video.png) no-repeat');
            $('#AVChatVideo').css('background-size', '50px 50px');
            $('#AVWindowSelfView').css('background-image', 'url(../Images/default.png)');
            //$('#AVWindowContactView').css('background-image', 'url(../Images/default.png)');
            video = false;
        }
        
    });

    $('#AVChatMute').click(function () {
        if (!muted) {
            //mute self-mic
            $('#AVChatMute').css('background', 'url(../Images/fabrikam_skypeControl_mic.png) no-repeat');
            $('#AVChatMute').css('background-size', '50px 50px');
            muted = true;
        } else {
            //unmute self-mic
            $('#AVChatMute').css('background', 'url(../Images/fabrikam_skypeControl_mic_muted.png) no-repeat');
            $('#AVChatMute').css('background-size', '50px 50px');
            muted = false;
        }
    });

    $('#AVChatClose').click(function () {
        endConversation();
        window.close();
    });
});