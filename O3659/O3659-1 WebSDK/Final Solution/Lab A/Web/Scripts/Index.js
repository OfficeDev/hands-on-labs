(function () {
    $(document).ready(function () {
       
        $('#skypeLogin').click(showSkypeLogin);
        $('#AnonChat').click(showSkypeLogin);
        if (/^#access_token=/.test(location.hash)) {
            console.log('authenticated');
            location.assign('Home.html?auto=1&ss=0' +
                '&cors=1' +
                '&client_id=' + config.clientId+
                '&origins=https://webdir.online.lync.com/autodiscover/autodiscoverservice.svc/root');
        }

    });
    function showSkypeLogin() {
        location.assign('https://login.microsoftonline.com/common/oauth2/authorize?response_type=token' +
                '&client_id=' + config.clientId+
                '&redirect_uri=' + config.redirect_uri+
                '&resource=https://webdir.online.lync.com');
    }

})();