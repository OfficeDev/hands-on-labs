var SkypeMessenger = (function () {
    var SkypeWebApp;
    var SkypeApi;

    var Initialize = function () {

        var dfd = jQuery.Deferred();

       //code goes here

        return dfd;
    }

    //this is a helper method used for loading components located in html/messenger folder
    var loadedModules = [];
    var Load = function(module, location){

        var dfd = jQuery.Deferred();
        var alreadyLoaded = false;

        var _module = {
            name: module,
            location: location
        }

        loadedModules.forEach(function (e) {
            if (e.name == module && e.location == location)
                alreadyLoaded = true;
        });

        $(location).children().hide();

        if (alreadyLoaded) {
            var c = '.' + module;
            $(location + ' ' + c).show();
            dfd.resolve(_module);
        } else {
            var url = "html/messenger/" + module + ".html";

            $.ajax({
                url: url,
                async: true,
                context: document.body,
                success: function (html) {
                    $(location).append(html);
                    dfd.resolve(_module);
                }
            });                

            loadedModules.push(_module);
        }

        return dfd.promise();
    }

    var LoadAzureSignIn = function () {
        //code goes here
    }

    var SignIn = function () {
        var dfd = jQuery.Deferred();

        //code goes here

        return dfd;
    }

    var SignInAnonymous = function (name, meetingUri) {
        var dfd = jQuery.Deferred();
       
        //code goes here

        return dfd;
    }

    var SignOut = function () {
        var dfd = jQuery.Deferred();

        //code goes here

        return dfd;
    }

    var ChatService = {
        CloseAllChats: function () {
            var client = SkypeMessenger.SkypeWebApp;

            var dfd = jQuery.Deferred();

            //code goes here        

            return dfd.promise();
        },

        CreateConversation: function (sip) {
            var client = SkypeMessenger.SkypeWebApp;
            var pSearch = client.personsAndGroupsManager.createPersonSearchQuery();
            var dfd = jQuery.Deferred();

            //code goes here

            return dfd.promise();
        }
    }

    return {
        Initialize: Initialize,
        Load: Load,
        SkypeWebApp: SkypeWebApp,
        SkypeApi: SkypeApi,
        SignIn: SignIn,
        LoadAzureSignIn: LoadAzureSignIn,
        ChatService: ChatService,
        SignOut: SignOut,
        SignInAnonymous: SignInAnonymous
    }
})();

$(function () {
});