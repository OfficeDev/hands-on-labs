# Create your first Office add-in with the Outlook JavaScript APIs

In this lab, you will use Visual Studio to create your first Outlook add-in using the Outlook JavaScript APIs. The add-in will allow the user to translate parts of a message they are composing into different languages.

## Prerequisites

1. You must have an Office 365 tenant to
   complete this lab. If you do not have one, the lab for **O3651-7 Setting up
   your Developer environment in Office 365** shows you how to obtain a trial.
1. You must have Visual Studio 2015 with Update 1 installed.
1. You must have the Microsoft Office Developer Tools for Visual Studio 2015 installed.
1. You must have Outlook 2016 installed.

## Exercise 1: Create a new Outlook add-in project

In this exercise you will create a new project using the Outlook add-in template.

1. Launch Visual Studio 2015 and select **New**, **Project**.
  1. Expand **Templates**, **Visual C#**, **Office/SharePoint** ,**Web add-ins**. Select **Outlook Add-in**. Name the project **Translator** and click **OK**.
  
    ![The new project dialog using the Outlook add-in template](./Images/create-project.PNG)
  
1. Run the app to verify it works.
  1. Press F5 to begin debugging.
  1. When prompted, enter the email address and password of your Office 365 account. Visual Studio will install the add-in for that user.
  
    ![The Connect to Exchange email account dialog](./Images/deploy-addin.PNG)
    
  1. With the app running, open Outlook 2016 and logon to the user's mailbox. You should see a **Display all propeties** button on the ribbon when you select or open a message.
  
    ![A message in Outlook 2016 with the add-in button on the ribbon](./Images/default-button.PNG)
    
Now that you've verified that the add-in is working, Exercise 1 is complete!

## Exercise 2: Add buttons to the new message ribbon

In this exercise you will add a button to do English-to-Spanish translation and a button to open a task pane, allowing the user to select start and end languages. 
  
1. Add the **Translator** button group to the new message ribbon.
  1. Open the **Translator/TranslatorManifest/Translator.xml** file.
  1. Locate the following line:
  
    ```xml
    <FunctionFile resid="functionFile" />
    ```
    
  1. Insert the following after that line:
  
    ```xml
    <!-- Message Compose -->
    <ExtensionPoint xsi:type="MessageComposeCommandSurface">
      <OfficeTab id="TabDefault">
        <Group id="msgComposeGroup">
          <Label resid="groupLabel"/>
          <!-- Add English to Spanish button here -->
          <!-- Add More Options button here -->
        </Group>
      </OfficeTab>
    </ExtensionPoint>
    ```
    
1. Add the **English to Spanish** button.
  1. Replace the `<!-- Add English to Spanish button here -->` line with the following:
    
    ```xml
    <Control xsi:type="Button" id="msgComposeEn-Es">
      <Label resid="englishSpanishLabel"/>
      <Supertip>
        <Title resid="englishSpanishTitle"/>
        <Description resid="englishSpanishDesc"/>
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon16"/>
        <bt:Image size="32" resid="icon32"/>
        <bt:Image size="80" resid="icon80"/>
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>translateEnglish2Spanish</FunctionName>
      </Action>
    </Control>
    ```
    
1. Add the **More Options** button.
  1. Replace the `<!-- Add More Options button here -->` line with the following:
  
    ```xml
    <Control xsi:type="Button" id="msgComposePaneButton">
      <Label resid="translatePaneButtonLabel"/>
      <Supertip>
        <Title resid="translatePaneButtonTitle"/>
        <Description resid="translatePaneButtonDesc"/>
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon16"/>
        <bt:Image size="32" resid="icon32"/>
        <bt:Image size="80" resid="icon80"/>
      </Icon>
      <Action xsi:type="ShowTaskpane">
        <SourceLocation resid="translatePaneUrl"/>
      </Action>
    </Control>
    ```
    
1. Add resources for the new buttons.
  1. Locate the `<bt:Urls>` element within the `<Resources>` element in **Translator/TranslatorManifest/Translator.xml**.
  1. Add the following element after the last `<bt:Url>` element:
    
    ```xml
    <bt:Url id="translatePaneUrl" DefaultValue="~remoteAppUrl/TranslatePane.html"/>
    ```
  
  1. Locate the `<bt:ShortStrings>` element within the `<Resources>` element.
  1. Change the `DefaultValue` of the `<bt:String>` element with an `id` attribute of `groupLabel` to `Translator`.
    
    ```xml
    <bt:String id="groupLabel" DefaultValue="Translator"/>
    ```
  1. Add the following elements after the last `<bt:String>` element inside the `<bt:ShortStrings>` element:
  
    ```xml
    <bt:String id="englishSpanishLabel" DefaultValue="English to Spanish"/>
    <bt:String id="englishSpanishTitle" DefaultValue="Translate English to Spanish"/>
    <bt:String id="translatePaneButtonLabel" DefaultValue="More Options"/>
    <bt:String id="translatePaneButtonTitle" DefaultValue="Choose to and from language"/>
    ```
    
  1. Locate the `<bt:LongStrings>` element within the `<Resources>` element.
  1. Add the following elements after the last `<bt:String>` element inside the `<bt:LongStrings>` element:
  
    ```xml
    <bt:String id="englishSpanishDesc" DefaultValue="Translates the selected text from English to Spanish"/>
    <bt:String id="translatePaneButtonDesc" DefaultValue="Opens a window allowing you to choose a to and from language for translation"/>
    ```
    
  1. When you've made all of those changes, the `<Resources>` section of your file should look like the following:
  
    ```xml
    <Resources>
      <bt:Images>
        <bt:Image id="icon16" DefaultValue="~remoteAppUrl/Images/icon16.png"/>
        <bt:Image id="icon32" DefaultValue="~remoteAppUrl/Images/icon32.png"/>
        <bt:Image id="icon80" DefaultValue="~remoteAppUrl/Images/icon80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="functionFile" DefaultValue="~remoteAppUrl/Functions/FunctionFile.html"/>
        <bt:Url id="messageReadTaskPaneUrl" DefaultValue="~remoteAppUrl/MessageRead.html"/>
        <bt:Url id="translatePaneUrl" DefaultValue="~remoteAppUrl/TranslatePane.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="groupLabel" DefaultValue="Translator"/>
        <bt:String id="customTabLabel"  DefaultValue="My Add-in Tab"/>
        <bt:String id="paneReadButtonLabel" DefaultValue="Display all properties"/>
        <bt:String id="paneReadSuperTipTitle" DefaultValue="Get all properties"/>
        <bt:String id="englishSpanishLabel" DefaultValue="English to Spanish"/>
        <bt:String id="englishSpanishTitle" DefaultValue="Translate English to Spanish"/>
        <bt:String id="translatePaneButtonLabel" DefaultValue="More Options"/>
        <bt:String id="translatePaneButtonTitle" DefaultValue="Choose to and from language"/> 
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="paneReadSuperTipDescription" DefaultValue="Opens a pane displaying all available properties. This is an example of a button that opens a task pane."/>
        <bt:String id="englishSpanishDesc" DefaultValue="Translates the selected text from English to Spanish"/>
        <bt:String id="translatePaneButtonDesc" DefaultValue="Opens a window allowing you to choose a to and from language for translation"/>
      </bt:LongStrings>
    </Resources>
    ```
    
1. Save your changes and press F5 to start debugging. Once the app starts, open Outlook 2016. (If it is still open from before, the add-in should refresh after a moment.) Create a new message. You should see the **English to Spanish** and **More Options** buttons on the ribbon:

  ![A new message in Outlook 2016 with the add-in buttons](./Images/compose-message.PNG)
  
Now that the buttons are showing up in Outlook, Exercise 2 is complete!
  
## Exercise 3: Add translation

In this exercise you will implement the functions to call the [Yandex Translate API](https://translate.yandex.com/developers) and replace selected text in the message that is being composed.

1. Obtain a free Yandex API key.
  1. Go to https://translate.yandex.com/developers in your browser.
  1. Under **Getting Started**, click the **Get a free API key** link.
  1. Register and get your API key. Copy this key, you will need it later.

1. Add the code to call the Yandex API and do the translation. 
  1. Expand the **TranslatorWeb** project in Visual Studio. Right-click the **Scripts** folder and choose **Add**, then **JavaScript** file. Name the file `translate` and click **OK**. Add the following code:

    ```javascript
    // Helper function to generate an API request
    // URL to the Yandex translator service
    function generateRequestUrl(sourcelang, targetlang, text) {
      // Split the selected data into individual lines
      var tempLines = text.split(/\r\n|\r|\n/g);
      var lines = [];

      // Add non-empty lines to the data to translate
      for (var i = 0; i < tempLines.length; i++)
        if (tempLines[i] != '')
          lines.push(tempLines[i]);

      // Add each line as a 'text' query parameter
      var encodedText = '';
      for (var i = 0; i < (lines.length) ; i++) {
        encodedText += '&text=' + encodeURI(lines[i].replace(/ /g, '+'));
      }

      // API Key for the yandex service
      // Get one at https://translate.yandex.com/developers
      var apiKey = 'PASTE YOUR YANDEX API KEY HERE';

      return 'https://translate.yandex.net/api/v1.5/tr.json/translate?key='
        + apiKey + '&lang=' + sourcelang + '-' + targetlang + encodedText;
    }

    function translate(sourcelang, targetlang, callback) {
      Office.context.mailbox.item.getSelectedDataAsync('text', function (ar) {
        // Make sure there is a selection
        if (ar === undefined || ar === null ||
            ar.value === undefined || ar.value === null ||
            ar.value.data === undefined || ar.value.data === null) {
          // Display an error message
          callback('No text selected! Please select text to translate and try again.');
          return;
        }

        try {
          // Generate the API call URL
          var requestUrl = generateRequestUrl(sourcelang, targetlang, ar.value.data);

          $.ajax({
            url: requestUrl,
            jsonp: 'callback',
            dataType: 'jsonp',
            success: function (response) {
              var translatedText = response.text;
              var textToWrite = '';

              // The response is an array of one or more translated lines.
              // Append them together with <br/> tags.
              for (var i = 0; i < translatedText.length; i++)
                textToWrite += translatedText[i] + '<br/>';

              // Replace the selected text with the translated version
              Office.context.mailbox.item.setSelectedDataAsync(textToWrite, { coercionType: 'html' }, function (asyncResult) {
                // Signal that we are done.
                callback();
              });
            }
          });
        }
        catch (err) {
          // Signal that we are done.
          callback(err.message);
        }
      });
    }
    ```
    
  1. Replace the `PASTE YOUR YANDEX API KEY HERE` text with the Yandex API key you obtained earlier.

1. Add a UI-less function for the **English to Spanish** button.
  1. Open the **TranslateWeb/Functions/FunctionFile.html** file and add a `<script>` tag for the `translate.js` file you just created. Be sure to add this **before** the tag for `FunctionFile.js`.

    ```html
    <script src="../Scripts/translate.js" type="text/javascript"></script>
    <script src="FunctionFile.js" type="text/javascript"></script>
    ```

  1. Open the **TranslateWeb/Functions/FunctionFile.js** file and add the following function.

    ```javascript
    function translateEnglish2Spanish(event) {
      translate('en', 'es', function(error) {
        if (error) {
          Office.context.mailbox.item.notificationMessages.addAsync('translateError', {
            type: 'errorMessage',
            message: error
          });
        }
        else {
          Office.context.mailbox.item.notificationMessages.addAsync('success', {
            type: 'informationalMessage',
            icon: 'icon16',
            message: 'Translated successfully',
            persistent: false
          });
        }
      });
      
      event.completed();
    }
    ```
1. Add a task pane for the **More Options** button.
  1. Right-click the **TranslatorWeb** project and select **Add**, then **HTML Page**. Name the page `TranslatePane` and click **OK**. Replace the contents of that file with the following.
  
    ```html
    <!DOCTYPE html>
    <html>
      <head>
        <meta charset="UTF-8" />
        <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
        <title></title>
        <script src="https://ajax.aspnetcdn.com/ajax/jQuery/jquery-1.9.1.min.js" type="text/javascript"></script>
        
        <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
        
        <link rel="stylesheet" href="https://appsforoffice.microsoft.com/fabric/2.1.0/fabric.min.css" />
        <link rel="stylesheet" href="https://appsforoffice.microsoft.com/fabric/2.1.0/fabric.components.min.css" />
        <link href="TranslatePane.css" rel="stylesheet" type="text/css" />
      
        <script src="Scripts/FabricUI/JQuery.Dropdown.js" type="text/javascript"></script>
         
        <script src="Scripts/translate.js" type="text/javascript"></script>
        <script src="TranslatePane.js" type="text/javascript"></script>
      </head>
      <body>
        <div id="content-header">
          <div class="padding">
            <h1 class="ms-font-xl ms-fontWeight-light ms-fontColor-white">Translate Text</h1>
          </div>
        </div>
        <div id="content-main">
          <div id="pending" class="ms-Overlay ms-Overlay--dark" style="text-align:center">
            <div class="ms-font-xxl" id="pending-message"></div>
          </div>
          <div id="translate-form" class="ms-Grid">
            <div class="ms-Grid-row">
              <div class="ms-Grid-col ms-u-sm12">
                <h2 class="ms-font-l ms-fontWeight-light">Select the text to translate in the body, choose starting and ending languages, then click <strong>Translate</strong>.</h2>
              </div>
            </div>
            <div class="ms-Grid-row">
              <div class="ms-Grid-col ms-u-sm12">
                <div class="ms-Dropdown" id="start-lang">
                  <label class="ms-Label">Starting language</label>
                  <i class="ms-Dropdown-caretDown ms-Icon ms-Icon--caretDown"></i>
                  <select class="ms-Dropdown-select">
                    <option value="none">Choose a language...</option>
                    <option id="start-English" value="en">English</option>
                    <option id="start-Spanish" value="es">Spanish</option>
                    <option id="start-French" value="fr">French</option>
                  </select>
                </div>
              </div>
              <div class="ms-Grid-col ms-u-sm12">
                <div class="ms-Dropdown" id="end-lang">
                  <label class="ms-Label">Ending language</label>
                  <i class="ms-Dropdown-caretDown ms-Icon ms-Icon--caretDown"></i>
                  <select class="ms-Dropdown-select">
                    <option>Choose a language...</option>
                    <option id="end-English" value="en">English</option>
                    <option id="end-Spanish" value="es">Spanish</option>
                    <option id="end-French" value="fr">French</option>
                  </select>
                </div>
              </div>
            </div>
            <div class="ms-Grid-row">
              <div class="ms-Grid-col ms-u-sm12">
                <button id="translateText" class="ms-Button">
                  <span class="ms-Button-label">Translate</span>
                  <span class="ms-Button-description">Sends the selected text to Yandex for translation</span>
                </button>
              </div>
            </div>
            <div class="ms-Grid-row">
              <div class="ms-Grid-col ms-u-sm12">
                <div id="error-box" class="ms-bgColor-error">
                  <div id="error-msg" class="ms-font-l ms-fontColor-error"></div>
                </div>
              </div>
            </div>
            <div class="ms-Grid-row">
              <div class="ms-Grid-col ms-u-sm12">
                <pre id="debug"></pre>
              </div>
            </div>
          </div>
        </div>
      </body>
    </html>
    ```
  
  1. Right-click the **TranslatorWeb** project and select **Add**, then **JavaScript file**. Name the file `TranslatePane` and click **OK**. Replace the contents of that file with the following.
  
    ```javascript
    (function () {
      'use strict';
      // The initialize function must be run each time a new page is loaded
      Office.initialize = function (reason) {
        $(document).ready(function () {
          $('#error-box').hide();
          $('#pending').hide();
          $('.ms-Dropdown').Dropdown();
          $('#translateText').click(doTranslate);
        });
      };

      function doTranslate() {
        $("#error-box").hide('fast');
        var startlang = $('#start-lang').children('.ms-Dropdown-title').text();
        var endlang = $('#end-lang').children('.ms-Dropdown-title').text();

        var startlangcode = $('#start-lang').find('#start-' + startlang.replace(/\s|\./g, ''));
        var endlangcode = $('#end-lang').find('#end-' + endlang.replace(/\s|\./g, ''));

        if (startlangcode.length > 0 && endlangcode.length > 0) {
          $('#pending-message').html('Working on your ' + startlang +
          ' to ' + endlang + ' translation request');
          $('#translate-form').hide('fast');
          $('#pending').show('fast');

          translate(startlangcode.val(), endlangcode.val(), function (error) {
            $('#pending').hide('fast');
            $('#translate-form').show('fast');
            if (error) {
              $('#error-msg').html('ERROR: ' + error);
              $('#error-box').show('fast');
            }
          });
        }
        else {
          $('#error-msg').html('Select languages!');
          $('#error-box').show('fast');
        }
      }
    })();
    ```
    
  1. Right-click the **TranslatorWeb** project and select **Add**, then **Style Sheet**. Name the file `TranslatePane` and click **OK**. Replace the contents of that file with the following.
  
    ```css
    ```
    
  1. Add the Fabric UI Dropdown plugin
    1. Download the [Jquery.Dropdown.js file](https://github.com/OfficeDev/Office-UI-Fabric/blob/master/src/components/Dropdown/Jquery.Dropdown.js) from GitHub.
    1. Move the file into the **TranslatorWeb/Scripts/FabricUI** folder in the project.
    1. Right-click the **TranslatorWeb/Scripts/FabricUI**, choose **Add**, then **Existing item**. Browse to the **Jquery.Dropdown.js** file in the **FabricUI** folder and click **Add**.
    
  