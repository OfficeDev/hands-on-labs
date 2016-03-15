//helper function used for website navigation
var navigation = (function () {
    var goTo = function(module) {

        var url = "html/" + module + ".html";

        $.ajax({
          url: url,
          async: true,
          context: document.body,
          success: function(html) {
            $(".content").html(html);
          }
        });
    };

    var setActiveTab = function(tab){
        $(".navigation-panel .tabs>div.active").removeClass("active");
        $(tab).addClass("active");
    }

    return {
        goTo: goTo,
        setActiveTab: setActiveTab
    }
})();

//helper function used to monitor promises
function monitorPromise(title, promise) {
    console.log(title, 'started');
    promise.then(function (res) {
        console.log(title, 'succeeded');
        console.log(res);
    }, function (err) {
        console.log(title, 'failed', err && err.stack || err);
        alert(title + ' failed:' + err);
    });
}

