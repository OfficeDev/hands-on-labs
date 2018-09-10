# New in the Excel JavaScript API: Creating Data Analysis Web Add-ins for Excel using Pivot Tables.

During this lab you will learn how to create pivot tables in Excel using the JavaScript API.

## Preparation

On this Hands-on Lab, you will use Script Lab to code and run your snippet. Script Lab is a Web Add-in built by Microsoft that you can help you test, build and share your code snippets rapid and conveniently. If you are not familiar with it, you can follow the instructions below.

1.  If you don’t see a “Script Lab” tab on your ribbon, please install it from the store. Otherwise continue to Exercise 1\.
2.  Click on the Insert Tab on the ribbon and then select “Get Add-ins” The Office Add-ins dialog will pop up.
3.  In the store tab, search for “Script Lab”, then click on “Add”
4.  At the end of the process you should see the “Script Lab Tab” on the Ribbon.

## [![Title: images/Image1536609758684.undefined](~WRS%7b9D5CB4B4-0889-4037-8EAB-D6C7C825BB81%7d_files/image001.png)](https://raw.githubusercontent.com/OfficeDev/hands-on-labs/master/images/Image1536609758684.undefined)

## Exercise 1

In this exercise, you'll insert sample data in the worksheet and then create a Pivot Table to summarize it.

### Step 1.1 Create a new Script Lab Snippet.

Click on the “Code” Button on the Script Lab Ribbon Tab. That will open a task pane like this one:

[![Title: images/Image1536609641084.undefined](~WRS%7b9D5CB4B4-0889-4037-8EAB-D6C7C825BB81%7d_files/image002.png)](https://raw.githubusercontent.com/OfficeDev/hands-on-labs/master/images/Image1536609641084.undefined)

### Step 1.2 : Setup HTML Page and point to the Office.js BETA end point.

In order to use the Pivot Table API you need to add a reference to the Office.js BETA library.  Click on the Libraries Tab and then make sure it points to:

 [https://appsforoffice.microsoft.com/lib/beta/hosted/office.js](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js)

Now let’s click on the HTML Tab and add 3 buttons to a) Insert sample data, b) create Pivot Table and c) add  Rows, columns and data to the Pivot Table.

```
<section class="setup ms-font-m">
```

Your HTML TAB should look like this:

```

```

Step 1.2 Create a chart Add a button in HTML page. Clicking the button will use range above to create a chart.

[https://appsforoffice.microsoft.com/lib/beta/hosted/office.js](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js)

## Exercise 2

Bind an event handler to the created chart.

Step 2.1 Register event.

```
chart.onActivated.add(chartActivated);
```

Step 2\. Change the visibility of button in event handler.

Use these commands to control the visibility of Chart. When user select the chart, show the button.

```
document.getElementById("customize").style.display = "block";
```

## Exercise 3

Add data labels to the chart and set properties:

```
chart.dataLabels.position = "Center";
```

## Next steps

Congratulations! You’ve completed the experiments! If you want to learn more about new comming APIs, please move to our [Github](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)

## Appendix

[Reference anwser](https://gist.github.com/79f15944334e208361bbb1aa7229ec3f)