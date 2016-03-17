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

function endConversation(conversation) {
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
                supportsSharing: false
            }
        });
        // whenever client.state changes, display its value
        SkypeWebApp.signInManager.state.changed(function (state) {
            console.log("Skype Client state changed to: " + state);
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

    $('#AVChatMute').click(function () {
        var muted = SkypeWebApp.conversationsManager.conversations(0).selfParticipant.audio.isMuted();
        if (muted) {
            //unmute self-mic
            $('#AVChatMute').css('background', 'url(../Images/fabrikam_skypeControl_mic_muted.png) no-repeat');
            $('#AVChatMute').css('background-size', '50px 50px');
            muted = false;
        } else {
            //mute self-mic
            $('#AVChatMute').css('background', 'url(../Images/fabrikam_skypeControl_mic.png) no-repeat');
            $('#AVChatMute').css('background-size', '50px 50px');
            muted = true;
        }
    });

    $('#AVChatVideo').click(function () {
        if (!escalated) {
            console.log('escalating to video');
            $('#AVChatVideo').css('background', 'url(../Images/video-off.png) no-repeat');
            $('#AVChatVideo').css('background-size', '50px 50px');
            $('#AVWindowSelfView').css('background-image', 'url(../Images/self-video.png)');
            escalated = true;
        } else {
            if (!video) {
                //enable self video
                $('#AVChatVideo').css('background', 'url(../Images/video-off.png) no-repeat');
                $('#AVChatVideo').css('background-size', '50px 50px');
                $('#AVWindowSelfView').css('background-image', 'url(../Images/self-video.png)');
                video = true;
            } else {
                //disable self video
                $('#AVChatVideo').css('background', 'url(../Images/video.png) no-repeat');
                $('#AVChatVideo').css('background-size', '50px 50px');
                $('#AVWindowSelfView').css('background-image', 'url(../Images/default.png)');
                video = false;
            }
        }
    });

    $('#AVChatClose').click(function () {
        var conversation = SkypeWebApp.conversationsManager.conversations(0);
        endConversation();
    });
});