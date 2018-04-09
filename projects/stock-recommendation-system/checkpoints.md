# Project 2 Checkpoints

This document provides a progression of sequential checkpoints to help focus your development efforts. Students, feel free but not obligated to follow these instructions.

Regardless of whether you would like your final project solution to process a single stock symbol at a time, or multiple stock symbols at a time, you are recommended to start with the simpler "Single Symbol Solution" first. Then you can optionally apply your learnings to the more complex "Multiple Symbol Solution".

If you haven't already reviewed the [Alpha Vantage API documentation](https://www.alphavantage.co/documentation/), take a few moments to do so. Familiarize yourself with the available API functions (e.g. `TIME_SERIES_INTRADAY`, `TIME_SERIES_DAILY`, `TIME_SERIES_DAILY_ADJUSTED`, etc.), as well as the available request parameters and example responses for each. Choose a request URL that you would like to start with. You are highly, highly, highly recommended to request CSV-formatted data.

## Single Symbol Solution

There are two possible approaches to developing this project.

You could develop new functionality in the same logical order as the program's execution, following the user input all the way through to the user output. Or you could start with easy, familiar functionality first (i.e. capturing inputs, reading cell values, and performing calculations) and work your way towards more difficult, unfamiliar functionality (i.e. issuing HTTP requests and writing CSV data).

If you choose the former approach, walk through each of the checkpoints in order. This will allow you to get to what you know you don't know as soon as possible, so you can stop stressing about it.

If you choose the latter approach, start with Checkpoint 1 and 2. Then download an example CSV file from the API and import it into your workbook. You can pretend it was produced as a result of Checkpoint 6, and you can immediately skip to Checkpoint 7 and 8. Then once you are done, go back and complete Checkpoints 3-6.

### Checkpoint 1: Capture User Inputs

  1. Decide how you would like to capture the user's stock symbol input. Cell Value? Input Box? Some other mechanism?
  2. Create a simple user interface with the desired input mechanism and an ActiveX Command Button to initiate the recommendation process.
  3. Write VBA code to capture the user input on button click and optionally display it in a message box as feedback to the user.

![an example user interface which prompts the user to input a stock symbol into cell E11 and then press a command button to initiate the recommendation process](example-interface.png)

### Checkpoint 2: Validate User Inputs

At this point it is not feasible to know whether or not the stock symbol input represents a real publicly-traded company. But do your best to perform some basic preliminary validations on the stock symbol, like making sure it is not numeric in nature, and making sure it resembles a stock symbol (i.e. less than around 6 characters, and no spaces). Hint: review the [documentation for string functions](https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/functions/string-functions) to find one that will count the number of characters in a given string.

If the input doesn't pass validations, display a message box to the user to help them fix the error, and stop the program's execution. Otherwise, optionally display a message box to let the user know the input has passed initial validations.

![a screenshot showing a numeric user input and a message box instructing the user to input a stock symbol](example-preliminary-input-validation.png)

> NOTE: Prevention of unnecessary requests helps save resources (i.e. computer processing time and costs) for both your program and the API server. This respectful best practice is most likely appreciated by the API developers :smiley:.

### Checkpoint 3: Compile Request URL

  1. Declare a new string variable called `RequestURL` and assign its value to be a hard-coded representation of your chosen request URL.
  2. Feel free to move on to the next checkpoints using your hard-coded request URL. Later on, when you are ready, compile the request URL dynamically using a variable to represent the user inputted stock symbol. You are further encouraged to abstract this dynamic URL-compilation logic into a function which accepts a stock symbol parameter and returns the corresponding request URL.

### Checkpoint 4: Request Stock Data

  1. If you have not already done so, configure Excel to use Microsoft WinHTTP Services.
  2. Use Microsoft WinHTTP Services to instantiate a new request object, and send the request.
  3. Optionally display the corresponding response in a message box, including the response status and the response text.

![a screenshot of a message box displaying the response status (200, OK) and response text (a CSV-formatted sting truncated because it is too large to fit in the message box)](example-response.png)

### Checkpoint 5: Handle Response Errors

  1. Try issuing a request using a stock symbol you know doesn't exist, maybe something like "OOPS" or "GGGG". See what kind of response comes back.
  2. Create a validation which determines whether the response is successful or not. It should parse the response text for some identifying characteristic that differentiates an error response from a successful response. Hint: perhaps use the `InStr()` function to check whether or not the response text includes the words "Error Message". If the response is an error, display a friendly error message and stop the program's execution. Otherwise allow the program to continue.

![a screenshot of a message box displaying the response status (200, OK) and response text (a JSON-formatted error message)](example-error-response.png)

![a screenshot of a message box displaying the a user-friendly validation message](example-response-validation.png)

### Checkpoint 6: Write Response to Sheet

  1. Manually create a new sheet to store the historical price data. Name it something like "outputs" or "stock-data".
  2. Parse the response text, which will be a CSV formatted string as long as you made the proper request. Use loops and arrays to write the corresponding values to the output sheet.

![a screenshot of a worksheet full of historical stock prices. it has columns for "timestamp", "open", "high", "low", "close", and "volume". And is has a row of corresponding values for each day.](example-output-sheet.png)

### Checkpoint 7: Perform Calculations

  1. Prepare to test your program's ability to perform calculations using data from the output sheet. To avoid making unnecessary requests during this time, take a moment to comment-out the part of your code which issues the HTTP request. This will allow you to continue to test your program without straining the resources of the API server. This is another friendly best practice that would be appreciated by the API developers. Don't worry, when you are done testing your calculations, you can un-comment the HTTP request logic.
  2. Programmatically read cell values from your output sheet to determine metrics such as the latest closing price and any other information relevant to making a final recommendation.
  3. Optionally display these calculations in a message box.

![a screenshot of a message box showing latest, min, max, and average closing price](example-calculations.png)

Hint: You may find it helpful to use [various `WorksheetFunction`s](https://msdn.microsoft.com/en-us/vba/excel-vba/articles/using-excel-worksheet-functions-in-visual-basic) like `SUM()`, `MIN()`, `MAX()`, etc.

### Checkpoint 8: Output Final Recommendation

  1. Use a message box or cell values to issue a recommendation conditionally based on the results of your calculations.

![a screenshot of a message box showing a recommendation for the user to buy the stock](example-recommendation.png)

## Multiple Symbol Solution

TBA - see if you can demonstrate your ability to achieve the single symbol solution first.
