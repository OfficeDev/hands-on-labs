# New in the Excel JavaScript API: Creating Data Analysis Web Add-ins for Excel using PivotTables.

During this lab you will learn the basics on how to create PivotTables with the Excel JavaScript API. This lab uses hard-coded sample data that will be generated on a separate worksheet, but the add-in could pull the information from numerous sources.

## Preparation

### Lab tips

**Tip #1**  
At the top of this pane are two tabs: **Instructions** and **Resources**. 

- The **Resources** tab contains the Windows 10 login credentials that you'll need for this lab's startup screen.

- The **Instructions** tab within this pane contains these lab instructions. Switch back to the **Instructions** tab after you've acquired the necessary login credentials from the **Resources** tab.

**Tip #2**  
This pane is resizeable. For an optimal viewing experience, you may wish to resize this pane to be wider than its default width.

**Tip #3**  
To copy/paste code from a code block into Script Lab, click the **[T]** button that appears next to the code block. *Avoid* clicking anywhere within a code block, as doing so will also copy/paste the code into the active application.

### Excel and ScriptLab

This lab is done with Script Lab. Script Lab is an add-in for developing and testing other add-ins. Within Excel, you can test, build, and share your solutions. Script Lab is available from the Add-ins Store.

Open Excel to begin. You will be prompted to sign in to Office. That is not necessary for the lab. Please dismiss the login window.

Select **Blank Workbook** to begin. 

![An image displaying where Script Lab is on the Excel ribbon.](images/image1.png)

## Step 1: Setup your sample in Script Lab

You'll prepare Script Lab to code and run your sample. We'll write an HTML front-end and code the program logic in TypeScript. Open a blank workbook in Excel to begin.

### Step 1.1: Create a new Script Lab snippet

Go to the **Script Lab** ribbon tab and select **Code**. Once Script Lab load the pane, select the `+` icon to create a new add-in snippet. That opens a task pane like this one:

![The "Script" tab of Script Lab.](images/image2.png)

### Step 1.2: Setup HTML Page

Select the **HTML** tab and add three buttons. These will insert sample data, create a PivotTable, and add hierarchies to the PivotTable. Replace the contents of the **HTML** tab with the following:

```html
<section class="setup ms-font-m">
    <h3>Set up</h3>
    <button id="setup" class="ms-Button">
        <span class="ms-Button-label">Add sample Data</span>
    </button>
</section>
<section class="samples ms-font-m">
    <h3>Create the PivotTable</h3>
    <button id="createPivot" class="ms-Button">
        <span class="ms-Button-label">Create</span>
    </button>
</section>
<section class="samples ms-font-m">
    <h3>Adjust the PivotTable</h3>
    <button id="adjustPivot" class="ms-Button">
        <span class="ms-Button-label">Add rows, columns, and data</span>
    </button>
</section>
```

Your **HTML** tab should look like this:
![Script Lab Tab](images/image3.png)

### Step 1.3 : Add  Event handlers for each button.

Now select the **Script** tab on the Script Lab task pane and add an event handler for each button.

Replace the existing code with the following:

```typescript
$("#setup").click(() => tryCatch(setup));
$("#createPivot").click(() => tryCatch(createPivot));
$("#adjustPivot").click(() => tryCatch(adjustPivot));

async function setup() {
    await Excel.run(async (context) => {
        // TODO-1: Fill a worksheet with sample data
    });
}

async function createPivot() {
    await Excel.run(async (context) => {
        // TODO-2: Create a PivotTable
    });
}

async function adjustPivot() {
    await Excel.run(async (context) => {
        // TODO-3: Add hierarachies to the PivotTable
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
```
## Step 2: Create your PivotTable 

Step 2.1: Insert sample data

Replace `TODO-1` with the following:

```typescript
        const sheetData = await OfficeHelpers.ExcelUtilities
            .forceCreateSheet(context.workbook, "Data");
        const sheetPivot = await OfficeHelpers.ExcelUtilities
            .forceCreateSheet(context.workbook, "Pivot");

        const data = [["Farm", "Type", "Classification", "Crates Sold at Farm", "Crates Sold Wholesale"],
        ["A Farms", "Lime", "Organic", 300, 2000],
        ["A Farms", "Lemon", "Organic", 250, 1800],
        ["A Farms", "Orange", "Organic", 200, 2200],
        ["B Farms", "Lime", "Conventional", 80, 1000],
        ["B Farms", "Lemon", "Conventional", 75, 1230],
        ["B Farms", "Orange", "Conventional", 25, 800],
        ["B Farms", "Orange", "Organic", 20, 500],
        ["B Farms", "Lemon", "Organic", 10, 770],
        ["B Farms", "Kiwi", "Conventional", 30, 300],
        ["B Farms", "Lime", "Organic", 50, 400],
        ["C Farms", "Apple", "Organic", 275, 220],
        ["C Farms", "Kiwi", "Organic", 200, 120],
        ["D Farms", "Apple", "Conventional", 100, 3000],
        ["D Farms", "Apple", "Organic", 80, 2800],
        ["E Farms", "Lime", "Conventional", 160, 2700],
        ["E Farms", "Orange", "Conventional", 180, 2000],
        ["E Farms", "Apple", "Conventional", 245, 2200],
        ["E Farms", "Kiwi", "Conventional", 200, 1500],
        ["F Farms", "Kiwi", "Organic", 100, 150],
        ["F Farms", "Lemon", "Conventional", 150, 270]];

        const range = sheetData.getRange("A1:E21");
        range.values = data;
        range.format.autofitColumns();

        sheetPivot.activate();

        await context.sync();
```

Step 2.2: Create the base PivotTable

Replace `TODO-2` with the following:

```typescript
        const rangeToAnalyze = context.workbook.worksheets.getItem("Data").getRange("A1:E21");
        const rangeToPlacePivot = context.workbook.worksheets.getItem("Pivot").getRange("A2");
        context.workbook.worksheets.getItem("Pivot").pivotTables.add("Farm Sales", rangeToAnalyze, rangeToPlacePivot);

        await context.sync();
```

Step 2.3: Add row, column, and data hierarchies

Replace `TODO-3` with the following:

```typescript
		const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

		//add row hierarchies
		pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
		pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));

		// add column hierarchies
		pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

		// add data hierarchies
		const myValue = pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold at Farm"));
		myValue.summarizeBy = Excel.AggregationFunction.sum;

		await context.sync();
```

## Step 3: Run your sample!

Select the **Run** button on the “Script Lab” ribbon tab.

![The "Run" icon for Script Lab.](images/image4.png)

You should see a Task Pane with the HTML you created in the previous step.

![The "Run" pane for Script Lab.](images/image5.png)

Click the buttons in top/down order to generate a PivotTable like this one:

![The resulting PivotTable.](images/image6.png)

## Next steps
Congratulations! You’ve completed the lab! Please explore the other samples in Script Lab. Click on the menu button and go to **Samples** to learn about the different ways you can use the Excel JavaScript APIs in your add-ins.

To learn more about PivotTables, visit [Work with PivotTables using the Excel JavaScript API](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-pivottables).

Visit our [JavaScript API reference](https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets) for more information about the latest Excel APIs.
