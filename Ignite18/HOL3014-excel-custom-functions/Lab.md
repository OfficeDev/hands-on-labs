# Create custom functions in Excel

Custom functions enable you to add new functions to Excel by defining those functions in JavaScript as part of an add-in. Users within Excel can access custom functions just as they would any other native function in Excel, such as `SUM()`. 

In this lab, you'll create a custom functions add-in project by using the Yo Office generator, try out a prebuilt custom function that performs a simple calculation, create a custom function that requests data from the web, and create a custom function that streams real-time data from the web.

## Lab tips

**Tip #1**  
At the top of this pane are two tabs: **Instructions** and **Resources**. 

- The **Resources** tab contains the two sets of credentials that you'll need for this lab:  
    - The Windows 10 login credentials that you'll use to login to the VM at the startup screen.
    - The O365 user credentials that you'll use to login to Excel Online.

- The **Instructions** tab within this pane contains these lab instructions. Switch back to the **Instructions** tab after you've acquired the necessary login credentials from the **Resources** tab.

**Tip #2**  
This pane is resizeable. For an optimal viewing experience, you may wish to resize this pane to be wider than its default width.

**Tip #3**  
To copy/paste code from a code block into VS Code or the command prompt, click the **[T]** button that appears next to the code block. *Avoid* clicking anywhere within a code block, as doing so will also copy/paste the code into the active application.

**Tip #4**  
Shortcuts for the applications that you'll use during this lab are located in the toolbar along the bottom of this screen.

## Exercise 1: Create a custom functions add-in project

You'll begin this tutorial by using the Yo Office generator to create the files that you need for your custom functions project.

1. Open a command prompt and run the following command to create your add-in project. By default, the project should be created in the **C:\Users\LabUser** folder.

    ```bash
    yo office
    ```

    ![Yo Office bash prompts for custom functions](images/yo-office-cfs-stock-ticker-2.png)

    Answer the prompts as follows:
    - Choose a project type: **Excel Custom Functions Add-in project (...)**
    - Choose a script type: **Javascript**
    - What do you want to name your add-in? **stock-ticker**
    
    After you complete the wizard, the generator will create the project and install supporting Node components. Wait until this process completes before you move on to next steps in this lab guide.
    
2. At the command prompt, run the following command to navigate to the project folder.

    ```bash
    cd stock-ticker
    ```

3. At the command prompt, run the following command to download required dependencies.
    
    ```bash
    npm install
    ```

4. At the command prompt, run the following command to start a local web server.
    
    ```bash
    npm run start-web
    ```

5. Open Edge (the web browser) and launch Excel Online by navigating to the following URL: **https://www.office.com/launch/excel**.

6. When the **Sign in** dialog appears, open the **Resources** tab within this lab manual and click the **[T]** button next to the Office 365 **Username** (to automatically type the username into the dialog's textbox), then choose **Next**.

7. When the **Enter password** dialog appears, open the **Resources** tab within this lab manual and click the **[T]** button next to the Office 365 **Password** (to automatically type the password into the dialog's textbox), then choose **Sign in**.

8. In the **Stay signed in?** dialog, choose **Yes**.

9. In Excel Online, choose **New blank workbook**. 

10. If the **Your OneDrive is not set up** dialog appears, choose **Go to OneDrive -->**. When you see confirmation that **Your OneDrive is ready**, return to Excel Online in the other browser tab.

11. In Excel Online, choose **New blank workbook**. 

12. Register your custom functions add-in in Excel Online by completing the following steps:

    - Select **Insert** > **Office Add-ins**. 
    - Choose **Upload My Add-in**. 
    - Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created (**C:\Users\LabUser\stock-ticker**).
    - Select the file **manifest.xml** and choose **Open**, then choose **Upload**.

## Exercise 2: Try out a prebuilt custom function

The custom functions project that you created by using the Yo Office generator contains some prebuilt custom functions, defined within the **src/customfunction.js** file. The **manifest.xml** file in the root directory of the project specifies that all custom functions belong to the **CONTOSO** namespace.

At this point, the prebuilt custom functions in your project are loaded and available within Excel Online. Try out the **ADD** custom function by completing the following steps in Excel Online:

1. Within a cell, type **=CONTOSO**. Notice that the autocomplete menu shows the list of all functions in the **CONTOSO** namespace.

2. In any cell of your workbook, type the text **=CONTOSO.ADD(10,200)** and press enter. 

The **ADD** custom function computes the sum of the two numbers that you specify as input parameters. Typing **=CONTOSO.ADD(10,200)** should produce the result **210** in the cell after you press enter.

_Note that when a call is made in Excel Online, you may see `#GETTING_DATA` appear in a cell. Once a value is returned, this notification should disappear._

## Exercise 3: Create a custom function that requests data from the web

What if you needed a function that could retrieve and display the price of a stock in real time? Custom functions are designed so that you can easily request data from the web asynchronously.
  
Complete the following steps to create a custom function named **stockPrice** that accepts a stock ticker (e.g., **MSFT**) and returns the price of that stock. This custom function uses the IEX Trading API, which is free and does not require authentication.

1. In this lab, you'll use Visual Studio Code as your code editor. Open a command prompt and navigate to the **C:\Users\LabUser\stock-ticker** folder, then run the following command to open Visual Studio Code:

    ```
    code .
    ```

1. In Visual Studio Code, open the file **src/customfunctions.js** and place your cursor on the blank line that immediately follows the end of the **increment** function. With your cursor on that blank line, press the **[T]** button next to the following JavaScript code block in this Lab Guide to automatically copy/paste the code into **src/customfunctions.js**, and then save the file.

   In this code, notice that the asynchronous function returns a JavaScript Promise with the data from the IEX Trading API. Asynchronous custom functions must either return a new Promise or use JavaScript's **async** / **await** syntax.
    
    ```javascript
    function stockPrice(ticker) {
        var url = "https://api.iextrading.com/1.0/stock/" + ticker + "/price";
        return fetch(url)
            .then(function(response) {
                return response.text();
            })
            .then(function(text) {
                return parseFloat(text);
            });
        // Note: in case of an error, the returned rejected Promise
        //    will be bubbled up to Excel to indicate an error.
    }
    CustomFunctionMappings.STOCKPRICE = stockPrice;
    ```
    
2. Before Excel can make this new function available to end-users, you must specify metadata that describes this function. In Visual Studio Code: 

    - Open the file **config/customfunctions.json** and place your cursor at the end of line 49, immediately following the curly brace that ends the object that defines the **INCREMENT** function. Add a comma to the end of this line and press **Enter** to add a new line.

    - With your cursor on the new line (line 50), press the **[T]** button next to the following JSON code block in this Lab Guide to automatically copy/paste the code into **config/customfunctions.json**, then save the file.

    ```json
    {
        "id": "STOCKPRICE",
        "name": "STOCKPRICE",
        "description": "Fetches current stock price",
        "helpUrl": "http://yourhelpurl.com",
        "result": {
            "type": "number",
            "dimensionality": "scalar"
        },  
        "parameters": [
            {
                "name": "ticker",
                "description": "stock ticker name",
                "type": "string",
                "dimensionality": "scalar"
            }
        ]
    }
    ```
   
3. You must reregister the add-in in Excel Online in order for the new function to be available to end-users. Reregister your custom functions add-in in Excel Online by completing the following steps:

    - Select **Insert** > **Office Add-ins**. 
    - Choose **Upload My Add-in**. 
    - Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created (**C:\Users\LabUser\stock-ticker**).
    - Select the file **manifest.xml** and choose **Open**, then choose **Upload**.

4. Now, let's try out the new function. In any cell of your workbook, type the text **=CONTOSO.STOCKPRICE("MSFT")** and press enter. You should see that the result in the cell is the current stock price for one share of Microsoft stock.

## Exercise 4: Create a custom function that streams real-time data from the web

The **stockPrice** function that you just created returns the price of a stock at a specific moment in time, but stock prices are always changing. Let's create a custom function that streams data from an API to get real-time updates on a stock price.

Complete the following steps to create a custom function named **stockPriceStream** that requests the price of the specified stock every 1000 milliseconds (provided that the previous request has completed). While the initial request is in-progress, you may see the placeholder value **#GETTING_DATA** the cell where the function is being called. When a value is returned by the function, **#GETTING_DATA** will be replaced by that value in the cell.

1. In Visual Studio Code, open the file **src/customfunctions.js** and then:

    - Place your cursor after the closing curly brace for the **stockPrice** function, and press **Enter** to add a new line after the **stockPrice** function. 
    
    - With your cursor on the new line, press the **[T]** button next to the following JavaScript code block in this Lab Guide to automatically copy/paste the code into **src/customfunctions.js**, and then save the file.
    
    ```javascript
    function stockPriceStream(ticker, handler) {
        var updateFrequency = 1000 /* milliseconds*/;
        var isPending = false;
        var timer = setInterval(function() {
            // If there is already a pending request, skip this iteration:
            if (isPending) {
                return;
            }
            var url = "https://api.iextrading.com/1.0/stock/" + ticker + "/price";
            isPending = true;
            fetch(url)
                .then(function(response) {
                    return response.text();
                })
                .then(function(text) {
                    handler.setResult(parseFloat(text));
                })
                .catch(function(error) {
                    handler.setResult(error);
                })
                .then(function() {
                    isPending = false;
                });
        }, updateFrequency);
        handler.onCanceled = () => {
            clearInterval(timer);
        };
    }
    CustomFunctionMappings.STOCKPRICESTREAM = stockPriceStream;
    ```

2. Before Excel can make this new function available to end-users, you must specify metadata that describes this function. In Visual Studio Code: 

    - Open the file **config/customfunctions.json** and place your cursor at the end of line 67, immediately following the curly brace that ends the object that defines the **STOCKPRICE** function. Add a comma to the end of this line and press **Enter** to add a new line.

    - With your cursor on the new line (line 68), press the **[T]** button next to the following JSON code block in this Lab Guide to automatically copy/paste the code into **config/customfunctions.json**, then save the file.
   
    ```json
    { 
        "id": "STOCKPRICESTREAM",
        "name": "STOCKPRICESTREAM",
        "description": "Streams real time stock price",
        "helpUrl": "http://yourhelpurl.com",
        "result": {
            "type": "number",
            "dimensionality": "scalar"
        },  
        "parameters": [
            {
                "name": "ticker",
                "description": "stock ticker name",
                "type": "string",
                "dimensionality": "scalar"
            }
        ],
        "options": {
            "stream": true,
            "cancelable": true
        }
    }
    ```

3. You must reregister the add-in in Excel Online in order for the new function to be available to end-users. Reregister your custom functions add-in in Excel Online by completing the following steps:

    - Select **Insert** > **Office Add-ins**. 
    - Choose **Upload My Add-in**. 
    - Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created (**C:\Users\LabUser\stock-ticker**).
    - Select the file **manifest.xml** and choose **Open**, then choose **Upload**.

4. Now, let's try out the new function. In any cell of your workbook, type the text **=CONTOSO.STOCKPRICESTREAM("MSFT")** and press enter. Provided that the stock market is open, you should see that the result in the cell is constantly updated to reflect the real-time price for one share of Microsoft stock.

## Next steps

Congratulations, you've successfully created custom functions in Excel! This hands-on-lab ends here, but for more information about Custom Functions, be sure to check out our [online docs](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-overview).

## Legal Information

Data provided free by [IEX](https://iextrading.com/developer/). View [IEX's Term of Use](https://iextrading.com/api-exhibit-a/). Microsoft's use of this API in this hands-on-lab is for educational purposes only. 
