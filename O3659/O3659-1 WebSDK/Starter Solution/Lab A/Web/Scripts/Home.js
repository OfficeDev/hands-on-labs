var client;
var apiManager;
var access_token;
var contactNo = 0;






$(function () {


    function monitor(title, promise) {
        console.log(title, 'started');
        promise.then(function (res) {
            console.log(title, 'succeeded', res);
        }, function (err) {
            console.log(title, 'failed', err && err.stack || err);
            alert(title + ' failed:' + err);
        });
    }

    function toggleChat(sip) {
        startConversation(sip);
    }

    function toggleContacts() {
        endConversation();
    }

    function displayContactCard(contactCardID) {
        //get user data from the Skype logins here
        var cardName = $('#' + contactCardID + ' .contactName').html();
        var cardPresence = $('#' + contactCardID + ' .contactPresence').html();
        var cardSIP = $('#' + contactCardID + ' .contactSIP').val();
        $('#ContactCard #ContactPresence').html(cardPresence);
        $('#ContactCard .contactCardSIP').val(cardSIP);
        $('#ContactCard').dialog({
            dialogClass: "no-close",
            title: cardName,
            draggable: false,
            resizable: false,
            position: { my: "right", at: "left", of: $('#' + contactCardID) },
            buttons: [
              {
                  icons: {
                      primary: "ui-icon-chat"
                  },
                  click: function () {
                      $(this).dialog("close");
                      //replace the contact list with the chat window. Add the call and video buttons there as well
                      $('#ChatContact').val(cardName); //we'd store actual info here, not just the name
                      toggleChat(cardSIP);
                  }
              },
              {
                  icons: {
                      primary: "ui-icon-call"
                  },
                  click: function () {
                      $(this).dialog("close");
                      var params = "sip=" + cardSIP + "&audioOnly=true";
                      launchAVWindow(params);
                      //pop open the call window for audio. You can end the call or initiate video from here
                  },
              },
              {
                  icons: {
                      primary: "ui-icon-video"
                  },
                  click: function () {
                      $(this).dialog("close");
                      var params = "sip=" + cardSIP + "&audioOnly=false";
                      launchAVWindow(params);
                      //pop open the call window for full audio and video. You can end the call or turn off video, etc.
                  }
              }
            ]
        });
    }

    function launchAVWindow(params) {
        window.open('http://secondonlineapp.azurewebsites.net/AVChat.html?' + params, 'mywindow', 'width=800,height=600');
    }


    
    
   

    //listener on contact cards
    $('#CCContainer').on('click', '.contact', function () {
        displayContactCard($(this).attr('id'));
    });

    //listener on chat window controls
    //audio
    $('#CCContainer').on('click', '#ChatCall', function () {
        launchAVWindow($('#ChatContact').val());
    });
    //video
    $('#CCContainer').on('click', '#ChatVideo', function () {
        launchAVWindow($('#ChatContact').val());
    });
    //end chat
    $('#CCContainer').on('click', '#ChatClose', function () {
        $('#ChatContact').val(''); //clear the current conversation data
        toggleContacts();
    });

    function showLoggedInUserData() {
        $('#LoadingGif').hide();
        $('#SkypeContent').show();
        $('#Content').show();
    }
});