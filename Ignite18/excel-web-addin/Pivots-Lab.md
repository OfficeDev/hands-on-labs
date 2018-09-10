# New in the Excel JavaScript API: Creating Data Analysis Web Add-ins for Excel using Pivot Tables.

During this lab you will learn how to create pivot tables in Excel using the JavaScript API.

## Preparation

On this Hands-on Lab, you will use Script Lab to code and run your snippet. Script Lab is a Web Add-in built by Microsoft that can use to easily  code, run and share your Office.js snippets rapid and conveniently. If you are not familiar with it, you can follow the instructions below.

1.  If you don’t see a “Script Lab” tab on your ribbon, please install it from the store. Otherwise continue to Exercise 1\.
2.  Click on the Insert Tab on the ribbon and then select “Get Add-ins” The Office Add-ins dialog will pop up.
3.  In the store tab, search for “Script Lab”, then click on “Add”
4.  At the end of the process you should see the “Script Lab Tab” on the Ribbon.

## ![Script Lab Tab](images/image1.png)

## Step 1 : Setup your snippet in Script Lab

In this exercise, you'll insert sample data in the worksheet and then create a Pivot Table to summarize it.

### Step 1.1 Create a new Script Lab Snippet.

Click on the “Code” Button on the Script Lab Ribbon Tab. That will open a task pane like this one:

![Script Lab Tab](images/image2.png)

### Step 1.2 : Setup HTML Page and point to the Office.js BETA end point.

In order to use the Pivot Table API you need to add a reference to the Office.js BETA library.  Click on the Libraries Tab and change the first line so that it points to:

 [https://appsforoffice.microsoft.com/lib/beta/hosted/office.js](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js)

Now let’s click on the HTML Tab and add 3 buttons to a) Insert sample data, b) create Pivot Table and c) add  Rows, columns and data to the Pivot Table.

```
<section class="setup ms-font-m">
```

Your HTML TAB should look like this:

```

```

### Step 1.3 : Add  Event handlers for each button.

Now click on the “Script” tab on the Script Lab task pane and add 3 event handlers for each button.

Your code should look like this:

$("#setup").click(() => tryCatch(setup));

$("#createPivot").click(() => tryCatch(createPivot));

$("#adjustPivot").click(() => tryCatch(adjustPivot));

async function setup() {

    await Excel.run(async (context) => {

        OfficeHelpers.UI.notify("setup");

        await context.sync();

    });

}

async function createPivot() {

    await Excel.run(async (context) => {

        OfficeHelpers.UI.notify("create");

        await context.sync();

    });

}

async function adjustPivot() {

    await Excel.run(async (context) => {

        OfficeHelpers.UI.notify("adjust");

        await context.sync();

    });

}

/** Default helper for invoking an action and handling errors. */

async function tryCatch(callback) {

    try {

        await callback();

    }

    catch (error) {

        OfficeHelpers.UI.notify(error);

        OfficeHelpers.Utilities.log(error);

    }

}

## Step  2: Create your Pivot Table.

Step 2.1 Add Code to the setup method (copy paste).

```
chart.onActivated.add(chartActivated);
```

Step 2.2 Add Code to add rows, columns and data hierarchies.

```
document.getElementById("customize").style.display = "block";
```

## Step 4: Run your sample!

Click on the RUN tab on “Script Lab”


![Script Lab Tab](images/image4.png)

You should see a Task Pane with the HTML you created in the previous step.


![Script Lab Tab](images/image5.png)

Click the buttons in tod/down order and you will see a pivot table like this one:


![Script Lab Tab](images/image5.png)

Son in this Lab you learned how to build a pivot table to summarize data.
