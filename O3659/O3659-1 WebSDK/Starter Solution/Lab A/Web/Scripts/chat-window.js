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