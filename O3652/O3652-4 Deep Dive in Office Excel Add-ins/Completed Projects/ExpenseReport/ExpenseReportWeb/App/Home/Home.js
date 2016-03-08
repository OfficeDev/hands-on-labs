/// <reference path="../App.js" />

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            // attach click handlers to the workbook
            // TODO-1
            $('#insertData').click(insertData);
            // TODO-2
            $('#sort').click(sort);
            // TODO-3
            $('#filter').click(filter);
            // TODO-4
            $('#report').click(report);
        });
    };

    function insertData() {
        Excel.run(function (ctx) {

            // Get the current worksheet
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();

            //Rename the current Worsheet to Data
            sheet.name = "Data";

            //Insert Data
            var range = sheet.getRange("A1:E11");
            range.values = [[
            "Date",
            "Merchant",
            "Category",
            "Sub-Category",
            "Amount"],
              [
                "01/12/2014",
                "WHOLE FOODS MARKET",
                "Merchandise & Supplies",
                "Groceries",
               "84.99"
              ],
              [
                "01/13/2014",
                "COSTCO GAS",
                "Transportation",
                "Fuel",
               "52.20"
              ],
              [
                "01/13/2014",
                "COSTCO WHOLESALE",
                "Merchandise & Supplies",
                "Wholesale Stores",
               "163.67"
              ],
              [
                "01/13/2014",
                "ITUNES",
                "Merchandise & Supplies",
                "Internet Purchase",
               "9.83"
              ],
              [
                "01/13/2014",
                "SMITH BROTHERS FARMS INC",
                "Merchandise & Supplies",
                "Groceries",
               "21.45"
              ],
              [
                "01/14/2014",
                "SHELL",
                "Transportation",
                "Fuel",
                "44.00"
              ],
              [
                "01/14/2014",
                "WHOLE FOODS MARKET",
                "Merchandise & Supplies",
                "Groceries",
               "17.98"
              ],
              [
                "01/15/2014",
                "BRIGHT EDUCATION SERVICES",
                "Other",
                "Education",
               "59.92"
              ],
              [
                "01/15/2014",
                "BRIGHT EDUCATION SERVICES",
                "Other",
                "Education",
               "59.92"
              ],
              [
                "01/17/2014",
                "SMITH BROTHERS FARMS INC-HQ",
                "Merchandise & Supplies",
                "Groceries",
               "21.45"
              ]];

            //Autofit row height and column width
            range.getEntireColumn().format.autofitColumns();
            range.getEntireRow().format.autofitRows();

            // Add a table
            var table = ctx.workbook.tables.add("Data!A1:E11", true);
            return ctx.sync().then(function () {
            });
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    function sort() {
        Excel.run(function (ctx) {

            var sheet = ctx.workbook.worksheets.getActiveWorksheet();

            // Only Sort the range that has data
            var sortRange = sheet.getRange("A1:E1").getEntireColumn().getUsedRange();
            // Apply sorting on the first column and in descending order
            sortRange.sort.apply([
            {
                key: 0,
                ascending: false,
            },
            ]);
            return ctx.sync().then(function () {
            })
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    function filter() {
        Excel.run(function (ctx) {
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            var table = sheet.tables.getItemAt(0);

            //Apply a value filter on the 4th column, which is sub-category. We want to focus on transactions in the category of Fuel and Education
            var filter = table.columns.getItemAt(3).filter;
            filter.applyValuesFilter(["Fuel", "Education"]);
            return ctx.sync().then(function () {
            })
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });

    }
    function report() {
        Excel.run(function (ctx) {
            //Add a new worksheet
            var sheet = ctx.workbook.worksheets.add("Summary");
            //Activate the worksheet
            sheet.activate();

            // Use Excel formulas to calculate the total spending based on categories
            var sumRange = sheet.getRange("A1:B6");
            sumRange.values = [['Category', 'Total'],
            ['Groceries', '=SUMIF( Data!D2:D100, "Groceries", Data!E2:E100 )'],
            ['Fuel', '=SUMIF( Data!D2:D100, "Fuel", Data!E2:E100 )'],
            ['Wholesale Store', '=SUMIF( Data!D2:D100, "Wholesale Stores", Data!E2:E100 )'],
            ['Internet Purchase', '=SUMIF( Data!D2:D100, "Internet Purchase", Data!E2:E100 )'],
            ['Education', '=SUMIF( Data!D2:D100, "Education", Data!E2:E100 )']];

            //Add a Table
            ctx.workbook.tables.add("Summary!A1:B6", true);

            // Add a pie chart
            var chartRange = sheet.getRange("A1:B6");
            var chart = ctx.workbook.worksheets.getItem("Summary").charts.add("Pie", chartRange);

            //Update the chart title
            chart.title.text = "Spending based on catagory";

            // Protect the report from editing
            sheet.protection.protect();

            return ctx.sync().then(function () {

            })
            .then(ctx.sync);
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

})();