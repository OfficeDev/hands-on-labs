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

})();