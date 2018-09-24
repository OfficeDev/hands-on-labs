# Build an Outlook add-in with Office JavaScript APIs

Outlook add-ins are web applications built using standard web technologies and loaded within the Outlook client. In this lab, you'll use Outlook JavaScript APIs to build an event-driven add-in that can validate room availability and capacity as an organizer creates an appointment.

![A screenshot of an Outlook meeting invitation with task pane](images/invite-two-recipients-johnson-pm.png)

In this lab:

- [Lab tips](#lab-tips)
- [Open Outlook and sign in](#open-outlook-and-sign-in)
- [Create the add-in project](#create-the-add-in-project)
- [Update the code](#update-the-code)
- [Prepare to test your add-in](#prepare-to-test-your-add-in)
- [Try it out](#try-it-out)

## Lab tips

**Tip #1**  
At the top of this pane are two tabs: **Instructions** and **Resources**. 

- The **Resources** tab contains the two sets of credentials that you'll need for this lab:  
    - The Windows 10 login credentials that you'll use to login to the VM at the startup screen.
    - The O365 user credentials that you'll use to login to Outlook.

- The **Instructions** tab within this pane contains these lab instructions. Switch back to the **Instructions** tab after you've acquired the necessary login credentials from the **Resources** tab.

**Tip #2**  
This pane is resizeable. For an optimal viewing experience, you may wish to resize this pane to be wider than its default width.

**Tip #3**  
To copy/paste code from a code block into VS Code or the command prompt, click the **[T]** button that appears next to the code block. *Avoid* clicking anywhere within a code block, as doing so will also copy/paste the code into the active application.

**Tip #4**  
Shortcuts for the applications that you'll use during this lab are located in the toolbar along the bottom of this screen.

## Open Outlook and sign in

A few minutes from now, you'll be testing the Outlook add-in that you've built. To ensure that Outlook is ready to go when you get to that point, complete the following steps now:

1. Open Outlook 2016. 

1. When the first Outlook dialog appears, open the **Resources** tab within this lab manual and click the **[T]** button next to the Office 365 **Username** (to automatically type the username into the dialog's textbox), then choose **Connect**. 

1. If prompted to choose account type, choose **Office 365**.

1. When the **Enter password** dialog appears, click the **[T]** button next to the Office 365 **Password** on the **Resources** tab (to automatically type the password into the dialog's textbox), then choose **Sign in**.

1. In the **Use the account everywhere on your device** dialog, accept the default settings and choose **Yes**.

1. In the **You're all set!** dialog, choose **Done**.

1. In the next dialog, de-select the **Set up Outlook Mobile on my phone, too** checkbox and choose **Done** (or **OK**). (If a browser window opens, close it and return to Outlook 2016.)

1. When Outlook launches, close the **Sign in to set up Office** dialog by clicking the **X** in the upper-right corner of the dialog.

1. If the **Make Office work smarter for you** dialog appears, select **No** and then choose **Accept**.

1. When Outlook opens, you'll see a **PRODUCT NOTICE** banner below the ribbon that indicates "Most of the features of Outlook have been disabled because it hasn't been activated." You can ignore this warning.

1. Close Outlook and then restart Outlook. The **PRODUCT NOTICE** banner should not be displayed after you've restarted Outlook.

Now that Outlook is configured, leave it running and move on to [creating the add-in project](#create-the-add-in-project)! 

## Create the add-in project

Complete the following steps to create the add-in project by using the **Yeoman generator for Office Add-ins**.

1. Open a command prompt and run the following command to create a new folder named `my-outlook-addin` at the root of the `C:\` drive. This is where you'll create the add-in project.

    ```
    mkdir C:\my-outlook-addin
    ```

1. At the command prompt, run the following command to navigate to your new folder.

    ```
    cd C:\my-outlook-addin
    ```

1. Use the **Yeoman generator for Office Add-ins** to create an Outlook Add-in project by running the following command from the command prompt, then answer the prompts as shown below. Please note that it may take several seconds before you see the first prompt.

    ```
    yo office
    ```

    - **Choose a project type:** `Office Add-in project using Jquery framework`
    - **Choose a script type:** `JavaScript`
    - **What do you want to name your add-in?:** `My Outlook Add-in`
    - **Which Office client application would you like to support?:** `Outlook`
    
    ![A screenshot of the prompts and answers for the Yeoman generator](images/yo-prompts.png)
    
    After you complete the wizard, the generator will create the project and install supporting Node components. Wait until this process completes before you move on to next steps in this lab guide.

## Update the code

At this point, the **Yeoman generator for Office Add-ins** has created a very basic add-in project that you can use as a starting point for building your Outlook add-in. Next, update the code as described in this section to customize the functionality of the add-in.

### Step 1: Open add-in project folder in Visual Studio Code

In this lab, you'll use Visual Studio Code as your code editor. At the same command prompt that you just used to [create the add-in project](#create-the-add-in-project), run the following command to open Visual Studio Code:

```
code .
```

In the **Explorer** pane of Visual Studio Code, expand the **My Outlook Add-in** folder to show the files for your add-in project.

### Step 2: Customize the manifest

An add-in's manifest file defines its settings and capabilities. In this step, you'll customize XML markup in the manifest file to specify metadata for the Room Validator add-in and add a button to the ribbon that an appointment organizer can use to launch the add-in task pane.

1. In Visual Studio Code, open the file **my-outlook-add-in-manifest.xml**. 

1. Delete all contents from the file and then put your cursor at the top of the empty file (line 1, column 1).

1. With your cursor located at the top of the empty file in Visual Studio code, press the **[T]** button next to the following XML code block in this Lab Guide to automatically copy/paste the code into **my-outlook-add-in-manifest.xml**, and then save the file. Notice that the `ExtensionPoint` element defines the button on the ribbon that will open the add-in's task pane. In this case, the button will appear on the ribbon only for an appointment organizer.

    ```xml
    <?xml version="1.0" encoding="UTF-8"?>
    <OfficeApp
            xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
            xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
            xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
            xmlns:mailappor="http://schemas.microsoft.com/office/mailappversionoverrides/1.0"
            xsi:type="MailApp">
        <Id>bac018a4-efdb-494b-aa48-dd7c9eec25c4</Id>
        <Version>1.0.0.0</Version>
        <ProviderName>Jane Doe</ProviderName>
        <DefaultLocale>en-US</DefaultLocale>
        <DisplayName DefaultValue="My Outlook Add-in" />
        <Description DefaultValue="Room Validator"/>
        <IconUrl DefaultValue="https://localhost:3000/assets/icon-32.png" />
        <HighResolutionIconUrl DefaultValue="https://localhost:3000/assets/hi-res-icon.png"/>
        <SupportUrl DefaultValue="https://localhost:3000" />
        <Hosts>
            <Host Name="Mailbox" />
        </Hosts>
        <Requirements>
            <Sets>
                <Set Name="Mailbox" MinVersion="1.1" />
            </Sets>
        </Requirements>
        <FormSettings>
            <Form xsi:type="ItemRead">
                <DesktopSettings>
                    <SourceLocation DefaultValue="https://localhost:3000/index.html"/>
                    <RequestedHeight>250</RequestedHeight>
                </DesktopSettings>
            </Form>
        </FormSettings>
        <Permissions>ReadWriteItem</Permissions>
        <Rule xsi:type="RuleCollection" Mode="Or">
            <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
        </Rule>
        <DisableEntityHighlighting>false</DisableEntityHighlighting>
        <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
            <Requirements>
                <bt:Sets DefaultMinVersion="1.3">
                    <bt:Set Name="Mailbox" />
                </bt:Sets>
            </Requirements>
            <Hosts>
                <Host xsi:type="MailHost">
                    <DesktopFormFactor>
                        <FunctionFile resid="functionFile" />
                        <!-- Button for Appointment Organizer -->
                        <ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
                            <OfficeTab id="TabDefault">
                                <Group id="apptComposeGroup">
                                    <Label resid="groupLabel" />
                                    <Control xsi:type="Button" id="apptComposeOpenPaneButton">
                                        <Label resid="apptComposeButtonLabel" />
                                        <Supertip>
                                            <Title resid="apptComposeSuperTipTitle" />
                                            <Description resid="apptComposeSuperTipDescription" />
                                        </Supertip>
                                        <Icon>
                                            <bt:Image size="16" resid="icon16" />
                                            <bt:Image size="32" resid="icon32" />
                                            <bt:Image size="80" resid="icon80" />
                                        </Icon>
                                        <Action xsi:type="ShowTaskpane">
                                            <SourceLocation resid="apptComposeTaskPaneUrl" />
                                        </Action>
                                    </Control>
                                </Group>
                            </OfficeTab>
                        </ExtensionPoint>
                    </DesktopFormFactor>
                </Host>
            </Hosts>
            <Resources>
                <bt:Images>
                    <bt:Image id="icon16" DefaultValue="https://localhost:3000/assets/icon-16.png"/>
                    <bt:Image id="icon32" DefaultValue="https://localhost:3000/assets/icon-32.png"/>
                    <bt:Image id="icon80" DefaultValue="https://localhost:3000/assets/icon-80.png"/>
                </bt:Images>
                <bt:Urls>
                    <bt:Url id="functionFile" DefaultValue="https://localhost:3000/function-file/function-file.html"/>
                    <bt:Url id="apptComposeTaskPaneUrl" DefaultValue="https://localhost:3000/index.html"/>
                </bt:Urls>
                <bt:ShortStrings>
                    <bt:String id="groupLabel" DefaultValue="My Add-in Group"/>
                    <bt:String id="customTabLabel"  DefaultValue="My Add-in Tab"/>
                    <bt:String id="apptComposeButtonLabel" DefaultValue="Room Validator"/>
                    <bt:String id="apptComposeSuperTipTitle" DefaultValue="Validate the choice of meeting room"/>
                </bt:ShortStrings>
                <bt:LongStrings>
                    <bt:String id="apptComposeSuperTipDescription" DefaultValue="Opens a pane which validates that the selected meeting room is available at the chosen time and can accommodate the number of invited attendees."/>
                </bt:LongStrings>
            </Resources>
            <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
                <Requirements>
                    <bt:Sets DefaultMinVersion="1.3">
                        <bt:Set Name="Mailbox" />
                    </bt:Sets>
                </Requirements>
                <Hosts>
                    <Host xsi:type="MailHost">
                        <DesktopFormFactor>
                            <FunctionFile resid="functionFile" />
                            <!-- Button for Appointment Organizer -->
                            <ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
                                <OfficeTab id="TabDefault">
                                    <Group id="apptComposeGroup">
                                        <Label resid="groupLabel" />
                                        <Control xsi:type="Button" id="apptComposeOpenPaneButton">
                                            <Label resid="apptComposeButtonLabel" />
                                            <Supertip>
                                                <Title resid="apptComposeSuperTipTitle" />
                                                <Description resid="apptComposeSuperTipDescription" />
                                            </Supertip>
                                            <Icon>
                                                <bt:Image size="16" resid="icon16" />
                                                <bt:Image size="32" resid="icon32" />
                                                <bt:Image size="80" resid="icon80" />
                                            </Icon>
                                            <Action xsi:type="ShowTaskpane">
                                                <SourceLocation resid="apptComposeTaskPaneUrl" />
                                            </Action>
                                        </Control>
                                    </Group>
                                </OfficeTab>
                            </ExtensionPoint>
                        </DesktopFormFactor>
                    </Host>
                </Hosts>
                <Resources>
                    <bt:Images>
                        <bt:Image id="icon16" DefaultValue="https://localhost:3000/assets/icon-16.png"/>
                        <bt:Image id="icon32" DefaultValue="https://localhost:3000/assets/icon-32.png"/>
                        <bt:Image id="icon80" DefaultValue="https://localhost:3000/assets/icon-80.png"/>
                    </bt:Images>
                    <bt:Urls>
                        <bt:Url id="functionFile" DefaultValue="https://localhost:3000/function-file/function-file.html"/>
                        <bt:Url id="apptComposeTaskPaneUrl" DefaultValue="https://localhost:3000/index.html"/>
                    </bt:Urls>
                    <bt:ShortStrings>
                        <bt:String id="groupLabel" DefaultValue="My Add-in Group"/>
                        <bt:String id="customTabLabel"  DefaultValue="My Add-in Tab"/>
                        <bt:String id="apptComposeButtonLabel" DefaultValue="Room Validator"/>
                        <bt:String id="apptComposeSuperTipTitle" DefaultValue="Validate the choice of meeting room"/>
                    </bt:ShortStrings>
                    <bt:LongStrings>
                        <bt:String id="apptComposeSuperTipDescription" DefaultValue="Opens a pane which validates that the selected meeting room is available at the chosen time and can accommodate the number of invited attendees."/>
                    </bt:LongStrings>
                </Resources>
            </VersionOverrides>
        </VersionOverrides>
    </OfficeApp>
    ```

### Step 3: Customize the HTML

The HTML markup in file **index.html** renders the user interface (UI) of the add-in task pane. In this step, you'll customize the HTML to create the task pane UI of the Room Validator add-in.

1. In Visual Studio Code, open the file **index.html**. 

1. Delete all contents from the file and then put your cursor at the top of the empty file (line 1, column 1).

1. With your cursor located at the top of the empty file in Visual Studio code, press the **[T]** button next to the following HTML code block in this Lab Guide to automatically copy/paste the code into **index.html**, and then save the file.

    ```html
    <!DOCTYPE html>
    <html>
        <head>
            <meta charset="UTF-8" />
            <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
            <meta name="viewport" content="width=device-width, initial-scale=1">
            <title>My Outlook Add-in</title>
            <!-- Office JavaScript API -->
            <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.debug.js"></script>
            <!-- LOCAL -->
            <link rel="stylesheet" href="node_modules/office-ui-fabric-js/dist/css/fabric.min.css" />
            <link rel="stylesheet" href="node_modules/office-ui-fabric-js/dist/css/fabric.components.css" />
            <!-- CDN -->
            <!-- For the Office UI Fabric, go to http://aka.ms/office-ui-fabric to learn more. -->
            <!--<link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.2.0/css/fabric.min.css" />-->
            <!--<link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.2.0/css/fabric.components.min.css" />-->
            <link href="app.css" rel="stylesheet" type="text/css" />
        </head>
        <body class="ms-font-m ms-welcome">
            <header class="ms-welcome__header ms-bgColor-themePrimary ms-u-fadeIn500">
                <h1 class="ms-fontSize-xxl ms-fontWeight-regular ms-fontColor-white">Room Validator</h1>
            </header>
            <main id="app-body" class="ms-welcome__main" style="display: none;">
                <div id="appointment-details">
                    <h2 class="ms-font-l ms-fontWeight-semibold ms-fontColor-neutralPrimaryAlt ms-u-slideUpIn20">Appointment details</h2>
                    <br />
                    <p class="ms-font-m ms-fontColor-neutralSecondaryAlt"><span class="ms-fontWeight-semibold">Total number of attendees:&#160;</span><label id="attendees-count"></label></p>
                    <p class="ms-font-m ms-fontColor-neutralSecondaryAlt"><span class="ms-fontWeight-semibold">Start time:&#160;</span><label id="start-time"></label></p>
                    <p class="ms-font-m ms-fontColor-neutralSecondaryAlt"><span class="ms-fontWeight-semibold">End time:&#160;</span><label id="end-time"></label></p>
                    <br />
                </div>
                <h2 class="ms-font-xl ms-fontWeight-semibold ms-fontColor-themeDark ms-u-slideUpIn20">Choose a room</h2>
                <br />
                <p class="ms-font-m">Choose a room from the list to see its capacity and availability. Press <b>Select</b> to specify the chosen room as the meeting location and see validation results.</p>
                <br />
                <select id="room" class="ms-font-m">
                    <option value="0n">-- Choose a room --</option>
                    <option value="2a">Conference Room Adams</option>
                    <option value="2p">Conference Room Carter</option>
                    <option value="4a">Conference Room Ford</option>
                    <option value="4p">Conference Room Johnson</option>
                    <option value="6a">Conference Room Lincoln</option>
                    <option value="6p">Conference Room Reagan</option>
                    <option value="8a">Conference Room Truman</option>
                    <option value="8p">Conference Room Wilson</option>
                </select>
                <ul class="ms-List ms-welcome__features ms-u-slideUpIn10">
                    <li class="ms-ListItem">
                        <i class="ms-Icon ms-Icon--People"></i>
                        <span class="ms-font-m ms-fontColor-neutralPrimary">Capacity:&#160;&#160;</span><label id="room-capacity"></label>
                    </li>
                    <li class="ms-ListItem">
                            <i class="ms-Icon ms-Icon--DateTime"></i>
                            <span class="ms-font-m ms-fontColor-neutralPrimary">Availability:&#160;&#160;</span><label id="room-availability"></label>
                    </li>
                </ul>
                <br />
                <button id="select" class="ms-Button ms-bgColor-themeDark">
                    <span class="ms-fontColor-themeDark ms-fontWeight-semibold">Select</span>
                </button>
                <br />
                <h2 class="ms-font-xl ms-fontWeight-semibold ms-fontColor-themeDark ms-u-slideUpIn20">Validation results</h2>
                <br />
                <label id="result-message"></label>
                <ul id="result-list" class="ms-List ms-welcome__features ms-u-slideUpIn10"></ul>
            </main>
            <script type="text/javascript" src="node_modules/jquery/dist/jquery.js"></script>
            <script type="text/javascript" src="node_modules/office-ui-fabric-js/dist/js/fabric.js"></script>
        </body>
    </html>
    ```

### Step 4: Customize the CSS

The CSS code in file **app.css** specifies the custom styles that are used to render the task pane UI. In this step, you'll customize the CSS to specify the styles that are used by **index.html** to render the task pane UI of the Room Validator add-in.

1. In Visual Studio Code, open the file **app.css**.

1. Delete all contents from the file and then put your cursor at the top of the empty file (line 1, column 1).

1. With your cursor located at the top of the empty file in Visual Studio code, press the **[T]** button next to the following CSS code block in this Lab Guide to automatically copy/paste the code into **app.css**, and then save the file.

    ```css
    html, body {
        width: 100%;
        height: 100%;
        margin: 0;
        padding: 0;
    }

    ul, p, h1, h2, h3, h4, h5, h6 {
        margin: 0;
        padding: 0;
    }

    .ms-welcome {
        position: relative;
        display: -webkit-flex;
        display: flex;
        -webkit-flex-direction: column;
        flex-direction: column;
        -webkit-flex-wrap: nowrap;
        flex-wrap: nowrap;
        min-height: 800px;
        min-width: 180px;
        overflow: auto;
        overflow-x: hidden;
    }

    .ms-welcome__header {
        min-height: 30px;
        padding: 0px;
        padding-bottom: 10px;
        display: -webkit-flex;
        display: flex;
        -webkit-flex-direction: column;
        flex-direction: column;
        -webkit-flex-wrap: nowrap;
        flex-wrap: nowrap;
        -webkit-align-items: center;
        align-items: center;
        -webkit-justify-content: flex-end;
        justify-content: flex-end;
    }

    .ms-welcome__header > h1 {
        margin-top: 10px;
        text-align: center;
    }

    .ms-welcome__main {
        display: -webkit-flex;
        display: flex;
        -webkit-flex-direction: column;
        flex-direction: column;
        -webkit-flex-wrap: nowrap;
        flex-wrap: nowrap;
        -webkit-align-items: center;
        align-items: left;
        -webkit-flex: 1 0 0;
        flex: 1 0 0;
        padding: 20px 20px;
    }

    .ms-welcome__main > h2 {
        width: 100%;
        text-align: left;
    }

    .ms-welcome__features.ms-List .ms-ListItem > .ms-Icon {
        font-size: 14pt;
    }

    .ms-welcome__features.ms-List .ms-ListItem > .ms-Icon {
        margin-right: 10px;
    }

    @media (min-width: 0) and (max-width: 350px) {
        .ms-welcome__features {
            width: 100%;
        }
    }
    ```

### Step 5: Customize the script

The content of file **src\index.js** specifies the script for the add-in. In this step, you'll specify code that enables the Room Validator add-in to validate room selection when an appointment organizer changes attendees or appointment time.

1. In Visual Studio Code, open the file **src\index.js**. 

1. Delete all contents from the file and then put your cursor at the top of the empty file (line 1, column 1).

1. With your cursor located at the top of the empty file in Visual Studio code, press the **[T]** button next to the following JavaScript code block in this Lab Guide to automatically copy/paste the code into **src\index.js**.

    ```javascript
    'use strict';
    (function () {
        var attendeeCount = 0;
        Office.initialize = (reason) => {
            $('#app-body').show();
            $('#appointment-details').hide();
            $(document).ready(function () {
                // specify functions for UI events
                $('#select').click(processRoomSelection);
                $('#room').on('change', processRoomChange);
                // set initial message
                $('#result-message').html('Choose a room from the list above and then press <b>Select</b> to see validation results.');
                // get initial values for appointment time and number of attendees
                getAppointmentTime();
                getNumberOfAttendees();
                // TODO-1: register event handler for the Office.EventType.AppointmentTimeChanged event
                // TODO-2: register event handler for the Office.EventType.RecipientsChanged event
            });
        };
        // TODO-3: processApptTimeChange()
        // TODO-4: processRecipientChange()
        function getAppointmentTime() {
            // get start time and end time of the appointment
            var promise = new Promise(function(resolve, reject) {
                Office.context.mailbox.item.start.getAsync(
                    function (asyncResult) {
                        if (asyncResult.status == Office.AsyncResultStatus.Failed){
                            console.log(asyncResult.error.message);
                        }
                        else {
                            $('#start-time').html(asyncResult.value.toLocaleString());
                            Office.context.mailbox.item.end.getAsync(
                                function (asyncResult) {
                                    if (asyncResult.status == Office.AsyncResultStatus.Failed){
                                        console.log(asyncResult.error.message);
                                    }
                                    else {
                                        $('#end-time').html(asyncResult.value.toLocaleString());
                                        resolve();
                                    }
                                });
                        }
                    });
            });
            return promise;
        }
        function getNumberOfAttendees() {
            // get the total number of attendees
            var promise = new Promise(function(resolve, reject) {
                Office.context.mailbox.item.requiredAttendees.getAsync(
                    function (asyncResult) {
                        if (asyncResult.status == Office.AsyncResultStatus.Failed){
                            console.log(asyncResult.error.message);
                        }
                        else {
                            var requiredAttendees = asyncResult.value;
                            attendeeCount = requiredAttendees.length;
                            $('#attendees-count').html(attendeeCount);
                            Office.context.mailbox.item.optionalAttendees.getAsync(
                                function (asyncResult) {
                                    if (asyncResult.status == Office.AsyncResultStatus.Failed){
                                        console.log(asyncResult.error.message);
                                    }
                                    else {
                                        var optionalAttendees = asyncResult.value;
                                        attendeeCount += optionalAttendees.length;
                                        $('#attendees-count').html(attendeeCount);
                                        resolve();
                                    }
                                }
                            );
                        }
                    }
                );
            });
            return promise;
        };
        function processRoomChange() {
            // get capacity of the room that's selected in the list and update UI element to display it
            getRoomCapacity();
            // get the availability period of the room that's selected in the list and update UI element to display it
            getRoomAvailabilityPeriod();
        }
        function processRoomSelection() {    
            validateRoomChoice();
            // set appointment location
            var roomLocation;
            if ($('#room').val().substring(0, 1) !== '0') {
                roomLocation = $('#room option:selected').text();
            } else {
                roomLocation = '';
            }
            Office.context.mailbox.item.location.setAsync(roomLocation, function (result) {
                if (result.status == Office.AsyncResultStatus.Failed) {
                    console.log(result.error.message);
                }
            });
        }
        function validateRoomChoice() {
            var roomCapacity = getRoomCapacity();
            var isRoomAvailable = getRoomAvailability();
            var isValid = true;
            getRoomAvailabilityPeriod();
            $('#result-list').empty();
            if ($('#room').val().substring(0, 1) !== '0') {
                // show validation result for room capacity 
                if (attendeeCount > roomCapacity) {
                    $('#result-list').append('<li class="ms-ListItem"><i class="ms-Icon ms-Icon--Cancel"></i>Room capacity is insufficient</li>');
                    isValid = false;
                } else {
                    $('#result-list').append('<li class="ms-ListItem"><i class="ms-Icon ms-Icon--Accept"></i>Room capacity is sufficient</li>');
                }
                // show validation result for room availability
                if (isRoomAvailable) {
                    $('#result-list').append('<li class="ms-ListItem"><i class="ms-Icon ms-Icon--Accept"></i>Room is available</li>');
                } else {
                    $('#result-list').append('<li class="ms-ListItem"><i class="ms-Icon ms-Icon--Cancel"></i>Room is unavailable</li>');
                    isValid = false;
                }
                // show message indicating validation results
                if (isValid === true) {
                    $('#result-message').html('The selected room (<b>' + $('#room option:selected').text() + '</b>) is valid.');
                } else {
                    $('#result-message').html('The selected room (<b>' + $('#room option:selected').text() + '</b>) is invalid. Fix issue(s) before sending this invitation.');
                }
            }
            else {
                $('#result-message').html('Choose a room from the list above and then press <b>Select</b> to see validation results.');
            }
        }
        function getRoomCapacity() {
            // Note: For simplicity, room capacity logic is hardcoded in this example code.
            // In a real-world implementation, room capacity data would likely be retrieved from a web service or database.
            // from value of selected list item, take first character (number = room capacity)
            var roomCapacity = $('#room').val().substring(0, 1);
            if (roomCapacity === '0') {
                $('#room-capacity').text('n/a');
            } else {
                $('#room-capacity').text(roomCapacity);
            }
            return roomCapacity;
        }
        function getRoomAvailabilityPeriod() {
            // Note: For simplicity, room availability logic is hardcoded in this example code.
            // In a real-world implementation, room availability data would likely be retrieved from a web service or database.
            // from value of selected DDL item, take second character
            //   a = available in the AM
            //   p = available in the PM
            var roomAvailability = $('#room').val().substring(1);
            if (roomAvailability === 'n') {
                $('#room-availability').text('n/a');
            } else if (roomAvailability === "a") {
                $('#room-availability').text('AM');
            } else {
                $('#room-availability').text('PM');
            }
        }
        function getRoomAvailability() {
            // Note: For simplicity, room availability logic is hardcoded in this example code.
            // In a real-world implementation, room availability data would likely be retrieved from a web service or database.
            // from value of selected DDL item, take second character 
            //   a = available in the AM
            //   p = available in the PM
            var roomAvailability = $('#room').val().substring(1);
            // determine whether start time and end time occur in the AM or PM
            var start = $('#start-time').text();
            var end = $('#end-time').text();
            var startPeriod = start.substring(start.length-2);
            var endPeriod = end.substring(end.length-2);
            // return availability result
            if ((roomAvailability === 'a' && startPeriod === 'AM' && endPeriod === 'AM') || (roomAvailability === 'p' && startPeriod === 'PM' && endPeriod === 'PM')) {
                return true;
            } else {
                return false;
            }
        }
    })();
    ```

1. Within **index.js**, find the comment labeled `TODO-1` and select that entire line. 

1. With the entire `TODO-1` comment line selected in Visual Studio code, press the **[T]** button next to the following JavaScript code block in this Lab Guide to automatically copy/paste the code into **src\index.js** (thereby replacing the `TODO-1` comment line). This code registers an event handler for the `Office.EventType.AppointmentTimeChanged` event. It specifies that when the appointment time changes, the `processApptTimeChange` function will run.

    ```javascript
    // register event handler for the Office.EventType.AppointmentTimeChanged event
    Office.context.mailbox.item.addHandlerAsync(Office.EventType.AppointmentTimeChanged, 
        processApptTimeChange,
        function (result) {
            if (result.status == Office.AsyncResultStatus.Failed) {
                console.log(result.error.message);
            }
        }
    );
    ```

1. Within **index.js**, find the comment labeled `TODO-2` and select that entire line. 

1. With the entire `TODO-2` comment line selected in Visual Studio code, press the **[T]** button next to the following JavaScript code block in this Lab Guide to automatically copy/paste the code into **src\index.js** (thereby replacing the `TODO-2` comment line). This code registers an event handler for the `Office.EventType.RecipientsChanged` event. It specifies that when recipients are added or removed, the `processRecipientChange` function will run.

    ```javascript
    // register event handler for the Office.EventType.RecipientsChanged event
    Office.context.mailbox.item.addHandlerAsync(Office.EventType.RecipientsChanged, 
        processRecipientChange,
        function (result) {
            if (result.status == Office.AsyncResultStatus.Failed) {
                console.log(result.error.message);
            }
        }
    );
    ```

1. Within **index.js**, find the comment labeled `TODO-3` and select that entire line. 

1. With the entire `TODO-3` comment line selected in Visual Studio code, press the **[T]** button next to the following JavaScript code block in this Lab Guide to automatically copy/paste the code into **src\index.js** (thereby replacing the `TODO-3` comment line). This is the event handler function that runs when the `Office.EventType.AppointmentTimeChanged` event occurs. The add-in retrieves the new appointment time, verifies whether the room is available at that time, and updates the task pane user interface to convey the result.

    ```javascript
    function processApptTimeChange(result) {
        // get new appointment time, then verify whether the selected room is available
        getAppointmentTime()
            .then(function() {
                validateRoomChoice();
            })
            .catch(function(e) {
                console.log(e);
            });
    };
    ```

1. Within **index.js**, find the comment labeled `TODO-4` and select that entire line. 

1. With the entire `TODO-4` comment line selected in Visual Studio code, press the **[T]** button next to the following JavaScript code block in this Lab Guide to automatically copy/paste the code into **src\index.js** (thereby replacing the `TODO-4` comment line). This is the event handler function that runs when the `Office.EventType.RecipientsChanged` event occurs. The add-in retrieves the new number of attendees, verifies whether the room capacity is sufficient for that number of attendees, and updates the task pane user interface to convey the result.

    ```javascript
    function processRecipientChange(result) {
        // get new number of attendees, then verify whether the room capacity is sufficient
        getNumberOfAttendees()
            .then(function() {
                validateRoomChoice();
            })
            .catch(function(e) {
                console.log(e);
            });
    }
    ```

1. Save the changes you've made to **index.js**. 

## Prepare to test your add-in

Now that the code changes are complete, you're *almost* ready to test your Room Validator add-in in Outlook. But first, prepare to test your add-in by completing the tasks described in this section.

### Start the web server

1. Open a command prompt and run the following command to navigate to the root directory of your project.

    ```
    cd C:\my-outlook-addin\My Outlook Add-in
    ```

1. At the command prompt in the root directory of your project, run the following command to start the application web server at `https://localhost:3000`.

    ```
    npm start
    ```

### Sideload the add-in's manifest in Outlook

Now that your add-in application is running on a local web server and your workstation trusts the local web server's self-signed certificate, you can upload the add-in's manifest file to Outlook. The manifest file defines your add-in's settings and capabilities, providing Outlook with the information it requires to run your add-in.

1. Open Outlook and select **Get Add-ins** from the ribbon of the **Home** tab.

    ![Outlook 2016 Home tab ribbon Get Add-ins button](images/home-ribbon-addins.png)

1. In the dialog window that opens, select **My add-ins**.

    ![Outlook 2016 My add-ins](images/my-add-ins.png)

1. In the **Custom add-ins** section at the bottom of the dialog window, select **Add a custom add-in** and then choose **Add from file**.

    ![Outlook 2016 My add-ins add custom add-in from file](images/my-add-ins-add-from-file.png)

1. In the **Choose File to Upload** dialog window, navigate to the folder **C:\my-outlook-addin\My Outlook Add-in**, select your add-in project's manifest file **my-outlook-add-in-manifest.xml**, and press **Open**. When the **Warning** dialog window appears, choose **Install**.

    ![Outlook 2016 Choose File to Upload dialog window](images/my-add-ins-upload-choose-file.png)

## Try it out

Now for the fun part -- it's time to try out the add-in that you've built. Use the Room Validator add-in to validate room availability and capacity as you create a meeting invitation.

1. Open Outlook 2016 and navigate to Calendar view.

1. In Calendar view, press **New Meeting** to create a new meeting.

1. In the ribbon of the meeting window, choose **Room Validator** to open the Room Validator task pane.

    ![A screenshot of the Outlook ribbon with the Room Validator button highlighted](images/ribbon-room-validator.png)

1. Specify meeting information as follows:
    
    - Add one recipient to the **To...** field. (You won't actually be sending this meeting invitation, so specify any recipient that you wish.)

    - Set **Start time** to **October 1, 2018** (10/1/2018) at **9:00 am**.

    - Set **End time** to **October 1, 2018** (10/1/2018) at **10:00 am**.

1. In the **Choose a room** section of the task pane, choose **Conference Room Carter** from the list and press **Select**. When you press **Select**, the conference room is specified in the appointment's **Location** field and the add-in determines whether the selected room is available and has capacity for the number of attendees specified in the invitation. 

    - The task pane shows that the capacity of the selected room is 2 and it's available only in the PM hours.
    - The meeting contains a total of 2 attendees (1 invitee + the organizer), which is within the capacity of the selected room, but occurs in the AM hours.
    - The **Validation results** section of the task pane conveys that room capacity is sufficient, but the room is unavailable at the specified time.

    ![A screenshot of an Outlook meeting invitation with task pane](images/invite-one-recipient-carter.png)

1. Add another recipient to the **To...** field of the meeting invitation. 

    - The meeting now contains a total of 3 attendees (2 invitees + the organizer), which exceeds the capacity of the selected room.
    - When you add the new recipient, the **Validation results** section of the task pane automatically updates to convey that room capacity is insufficient.  

    ![A screenshot of an Outlook meeting invitation with task pane](images/invite-two-recipients-carter.png)

1. Let's find a room with capacity large enough to accommodate the 3 attendees. In the **Choose a room** section of the task pane, choose **Conference Room Johnson** from the list and press **Select**. When you press **Select**, the conference room is specified in the appointment's **Location** field and the add-in determines whether the selected room is available and has capacity for the number of attendees specified in the invitation. 

    - The task pane shows that the capacity of the selected room is 4 and it's available only in the PM hours.
    - The meeting still contains a total of 3 attendees (2 invitees + the organizer), which is within the capacity of the selected room, but occurs in the AM hours.
    - The **Validation results** section of the task pane conveys that room capacity is sufficient, but that the room is unavailable at the specified time.
 
    ![A screenshot of an Outlook meeting invitation with task pane](images/invite-two-recipients-johnson-am.png)

1. Since the selected conference room is only available in the PM hours, change the **Start time** of the meeting to **2:00pm**.
 
    - The meeting still contains a total of 3 attendees (2 invitees + the organizer), but now occurs in the PM hours.
    - When you change the time of the meeting, the **Validation results** section of the task pane automatically updates to convey the room is available at the specified time.

    ![A screenshot of an Outlook meeting invitation with task pane](images/invite-two-recipients-johnson-pm.png)

## Congratulations!

Congratulations, you've successfully created an Outlook add-in! To learn more about creating Outlook add-ins, check out the Outlook add-ins developer documentation at [https://aka.ms/outlook-add-ins-docs](https://aka.ms/outlook-add-ins-docs).
