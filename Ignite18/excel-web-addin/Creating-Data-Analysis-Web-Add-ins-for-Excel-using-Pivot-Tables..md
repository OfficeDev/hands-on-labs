# Using new JS Chart and Events API to build interactive data visualization
This hands on lab will show you a simple introduction of creating an interactive add-in when user moves the selecting focus to a chart, and customizing the chart elements

## Preparation
This expriment can be done within ScriptLab. ScriptLab is a swiss knife for Excel JS APIs that you can test, build and share your solution fast and convenient.If you are not familliar with it, you can follow the instructions below. 

1. In your Excel, shift to "Insert" tab and click "My Add-ins" button. 
2. In the store tab, you'll see scriplab on the top of the list. Or you can search "Script Lab".
3. Click add and lauch from ribbon.
4. under the Libraries tab, change the first URL to
https://appsforoffice.microsoft.com/lib/beta/hosted/office.js

## Exercise 1
In this exercise, you'll create an event handler binded to a chart. Once the chart is activated, the "Customize" button will be enabled.

***Step 1.1*** Insert sample data.
Let's insert the following data into a worksheet.

```
["State", "2013", "2014", "2015", "2016", "2017"]
["California", 139, 304, 483, 785, 1308],
["Florida", 170, 366, 307, 708, 837],
["Hawaii", 158, 289, 387, 879, 735],
["South Carolina", 153, 251, 311, 413, 432],
["West Verginia", 620, 632, 654, 674, 684],
["Texas", 399, 446, 914, 953, 1312],
["Arizona", 126, 234, 364, 483, 594]
```

***Step 1.2*** Create a chart
Add a button in HTML page. Clicking the button will use range above to create a chart.

## Exercise 2
Bind an event handler to the created chart. 

***Step 2.1*** Register event.

```
chart.onActivated.add(chartActivated);
chart.onDeactivated.add(chartDeactivated);
```

***Step 2.*** Change the visibility of button in event handler.

Use these commands to control the visibility of Chart. When user select the chart, show the button.

```
document.getElementById("customize").style.display = "block";
document.getElementById("customize").style.display = "none";
```

## Exercise 3
Add data labels to the chart and set properties:

```
chart.dataLabels.position = "Center";
chart.dataLabels.separator = "\n";
chart.dataLabels.format.font.color = "#000000";
```

## Next steps
Congratulations! You’ve completed the experiments! If you want to learn more about new comming APIs, please move to our [Github](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)


## Appendix
[Reference anwser](https://gist.github.com/79f15944334e208361bbb1aa7229ec3f)
