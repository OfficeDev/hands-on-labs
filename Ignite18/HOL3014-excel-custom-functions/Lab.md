
# Create a streaming Excel custom function
In this lab, you will learn how to create custom functions which perform a simple calculation, request data from the web, and stream real-time data from the web.

## Exercise 1: Create your add-in project
You’ll begin this tutorial by using the Yo Office generator, which will automatically populate the files you need for your project.

1. In your command line interface, create a scaffold of your project (by default, this should be in your `C:\Users\LabUser` folder):
    
    ```bash
    yo office
    ```
    
    ![Yo Office bash prompts for custom functions](images/yo-office-excel-cfs-stock-ticker.PNG)
    
    Answer the prompts as directed below:  
    - Choose a project type: `Excel Custom Functions Add-in project (September 2018 Preview Refresh: Requires the Insider channel for Excel)`
    - Choose a script type: `Javascript`
    - What do you want to name your add-in? `stock-ticker`
    
    After you complete the wizard, the generator will create the project files.

    
2. Next, navigate to the root folder in your project using your command line interface. Run the following code:

    ```bash
    cd stock-ticker
    ```
    Run the following command to ensure your dependencies get downloaded
    
    ```bash
    npm install
    ```
    Run the following code to start a local web server: 
    
    ```bash
    npm run start-web
    ```
    
3. Open a web browser and copy and paste in the following URL: **`https://www.office.com/launch/excel`** to launch Excel Online. 
3. Sign in with your demo credentials, and open a new workbook. If you get an error about your OneDrive not being setup, click **GoTo OneDrive** to set it up, and then go back an open a new workbook.
4. Select **Insert > Add-ins**. Choose **Manage My Add-ins** and select **Upload My Add-in**. Click "Browse..." for your manifest file (`C:\Users\LabUser\stock-ticker\manifest.xml`), then click Open, select **Upload**.

Now the custom functions in your file will be loaded and ready to use. There are several pre-built functions for you in the Yo Office project. All are attached to a namespace called CONTOSO which is defined in the XML manifest file. Once you start typing `=CONTOSO` in a cell, the list of available functions will appear.

Let's call `=CONTOSO.ADD()`. This function adds any two numbers you provide as arguments. In any cell, type `=CONTOSO.ADD(1,2)`. It should deliver the answer 3.

_Note that when a call is made in Excel Online, you may see `#GETTING_DATA` appear in a cell. Once a value is returned, this notification should disappear._

## Exercise 2: Create your own custom function
What if you wanted a function which could fetch and display the current price of Microsoft stock? Custom functions are designed so you can easily make requests for data from the web asynchronously.
  
Complete the following steps to create a custom function named STOCKPRICE that accepts a stock ticker (e.g., "MSFT") and returns the price of that stock. You'll leverage the IEX Trading API, which is free and does not require authentication.

1. Open Visual Studio Code, and open up the `stock-ticker` folder.
1. Copy and paste the function below and add it to **./src/customfunctions.js**. 

   You'll notice in this code that your asychronous function returns a JavaScript Promise with the data from the IEX Trading API. Asynchronous functions require you to either return a new Promise or use JavaScript's async/await syntax.
    
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
    
2. In order for Excel to properly run this function, you must also add some metadata to the **./config/customfunctions.json** file.
 You'll notice that this JSON file describes the function, listing the types and dimensionality of the results and parameters.

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
   
3. You need to re-upload your manifest for this function to be useable.  In Excel Online, select **Insert > Add-ins**. Choose **Manage My Add-ins** and select **Upload My Add-in**. Browse for your manifest file, then select **Upload**.

4. In any cell of your workbook, enter the function `=CONTOSO.STOCKPRICE("MSFT")`. It should show you the current stock price for one share of Microsoft stock.

## Exercise 3: Create a streaming custom function
The previous function returned the stock price for Microsoft at a particular moment in time, but stock prices are always changing. With custom functions, it is possible to “stream” data from an API to get updates on stock prices in real time.  

To do this, you’ll create a new function, `=CONTOSO.STOCKPRICESTREAM`. It makes a request for updated data every 1000 milliseconds. When a call is made, you may see `#GETTING_DATA` appear in a cell. Once a value is returned, this notification should disappear.

1. Copy and paste the code below into **./src/customfunctions.js**.
    
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

2. Next, add to the **./config/customfunctions.json** file with the code below.
   
   You'll notice that this JSON file is very similar to the previous function's JSON file, but that a new section has been added for       "options." Because this function is streaming, you must specify this as true in the JSON. 
   
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

3. Again, upload your manifest file to re-register your changes to the files.  In Excel Online, select **Insert > Add-ins**. Choose **Manage My Add-ins** and select **Upload My Add-in**. Browse for your manifest file, then select **Upload**.

4. In any cell in your workbook, run the function `=CONTOSO.STOCKPRICESTREAM("MSFT")`. You should the price of Microsoft stock - which will be update in real time right in your workbook. 

## Next steps
Congratulations, you’ve completed the custom functions add-in tutorial! This hands-on-lab ends here, but be sure to check out our online docs to learn more about custom functions.

- Learn more about [custom functions on Microsoft Docs](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-overview)

## Legal Information
Data provided free by [IEX](https://iextrading.com/developer/). View [IEX's Term of Use](https://iextrading.com/api-exhibit-a/). Microsoft's use of this API in this hands-on-lab is for educational purposes only. 
