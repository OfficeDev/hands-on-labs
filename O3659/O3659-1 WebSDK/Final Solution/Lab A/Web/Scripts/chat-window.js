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
    var sip = getUrlVars()["sip"];
    var audioOnly = getUrlVars()["audioOnly"];
    var video = !(audioOnly == "true");
    var escalated = false;
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
            subscribeToChatEvents();
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
        $('#AVChatVideo').css('background', 'url(../Images/video-off.png) no-repeat');
        $('#AVChatVideo').css('background-size', '50px 50px');
        $('#AVWindowSelfView').css('background-image', 'url(../Images/self-video.png)');
    }


    function startConversation(sip) {
        var person;
        GetContactFromName(sip).then(function (results) {
            results.forEach(function (result) {
                person = result.result;
            });
            console.log('Creating conversation');
            var conversation = SkypeWebApp.conversationsManager.createConversation();
            console.log('Creating participant');
            var convParticipant = conversation.createParticipant(person)
            console.log('Participant: ' + person.displayName());
            console.log('Subscribing to participant');
            subscribeToParticipant(convParticipant);
            console.log('Adding participant to conversation');
            conversation.participants.add(convParticipant);
            console.log('Adding conversation to conversationManager');
            SkypeWebApp.conversationsManager.conversations.add(conversation);
            console.log('wait_1 (5 seconds) beginning');
            setTimeout(function () {
                console.log('wait_1 finished');
                if (audioOnly == "true") {
                    conversation.audioService.start();
                } else {
                    conversation.videoService.start();
                }
            }, 5000);
        });
    }

    function GetContactFromName(contactSIP) {
        var query = SkypeWebApp.personsAndGroupsManager.createPersonSearchQuery();
        query.text(contactSIP);
        query.limit(1);
        return query.getMore();
    }

    function subscribeToChatEvents() {
        SkypeWebApp.conversationsManager.conversations.added(function (conversation) {
            conversation.selfParticipant.audio.state.changed(function (newState, reason, oldState) {
                if (newState == 'Notified') {
                    console.log("Audio notified");
                    conversation.audioService.accept();
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
                    conversation.videoService.accept();
                    console.log("Video notified");
                }
                else if (newState == 'Connected') {
                    console.log("Connected to Video service");
                    console.log('wait_2 (5 seconds) beginning');
                    setTimeout(function () {
                        console.log('wait_2 finished');
                        conversation.selfParticipant.video.channels(0).stream.source.sink.container.set(document.getElementById("render-self-window"));
                        conversation.selfParticipant.video.channels(0).isStarted.set(true);
                        escalated = true;
                    }, 5000);
                }
                else if (newState == "Disconnected") {
                    console.log("Disconnected from video service");
                }
            });
        });
    }

    function subscribeToParticipant(participant) {
        participant.video.state.changed(function (newState, reason, oldState) {
            console.log('participant video state changed');

            if (newState == 'Connected') {
                console.log("The remote participant video has connected");
                participant.video.channels(0).stream.source.sink.container.set(document.getElementById("render-participant-window"));
                participant.video.channels(0).isStarted.set(true);
                console.log("The remote participant sink set has completed");
            }
        });
    }

    $('#AVChatMute').click(function () {
        var muted = SkypeWebApp.conversationsManager.conversations(0).selfParticipant.audio.isMuted();
        if (muted) {
            //unmyte self-mic
            $('#AVChatMute').css('background', 'url(../Images/fabrikam_skypeControl_mic_muted.png) no-repeat');
            $('#AVChatMute').css('background-size', '50px 50px');
            SkypeWebApp.conversationsManager.conversations(0).selfParticipant.audio.isMuted.set(false);
            console.log('enabling self audio');
        } else {
            //mute self-mic
            $('#AVChatMute').css('background', 'url(../Images/fabrikam_skypeControl_mic.png) no-repeat');
            $('#AVChatMute').css('background-size', '50px 50px');
            SkypeWebApp.conversationsManager.conversations(0).selfParticipant.audio.isMuted.set(true);
            console.log('muting self audio');
        }
    });
    $('#AVChatVideo').click(function () {
        if (!escalated) {
            SkypeWebApp.conversationsManager.conversations(0).videoService.start();
            console.log('escalating to video');
            $('#AVChatVideo').css('background', 'url(../Images/video-off.png) no-repeat');
            $('#AVChatVideo').css('background-size', '50px 50px');
            $('#AVWindowSelfView').css('background-image', 'url(../Images/self-video.png)');
            escalated = true;
        } else {
            if (!video) {
                //enable self video
                console.log('re-enabling self video');
                SkypeWebApp.conversationsManager.conversations(0).selfParticipant.video.channels(0).isStarted.set(true);
                $('#AVChatVideo').css('background', 'url(../Images/video-off.png) no-repeat');
                $('#AVChatVideo').css('background-size', '50px 50px');
                $('#AVWindowSelfView').css('background-image', 'url(../Images/self-video.png)');
                video = true;
            } else {
                //disable self video
                console.log('disabling self video');
                SkypeWebApp.conversationsManager.conversations(0).selfParticipant.video.channels(0).isStarted.set(false);
                $('#AVChatVideo').css('background', 'url(../Images/video.png) no-repeat');
                $('#AVChatVideo').css('background-size', '50px 50px');
                $('#AVWindowSelfView').css('background-image', 'url(../Images/default.png)');
                video = false;
            }
        }
    });
    $('#AVChatClose').click(function () {
        var conversation = SkypeWebApp.conversationsManager.conversations(0);
        endConversation(conversation);
    });

    function endConversation(conversation) {
        conversation.videoService.stop();
        conversation.audioService.stop();
        conversation.chatService.stop();
        conversation.leave();
        SkypeWebApp.conversationsManager.conversations.remove(conversation);
        location.assign('/Home.html');
    }
});