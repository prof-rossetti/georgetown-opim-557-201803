# Stock Recommendation System

Assume you own and operate a financial planning business which helps customers make investment decisions.

Your objective is to build yourself a tool to automate the process of providing your clients with stock trading recommendations.

Specifically, the system should accept one or more stock symbols as information inputs, and should provide a recommendation as to whether or not the client should purchase the given stock(s).

## Learning Objectives

  1. Design and build a tool to automate manual efforts and aid a decision-making process.
  2. Find practical applications for learning new programming concepts, primarily requesting and processing API data from the Internet.

## Instructions

Create a new macro-enabled workbook named **`netid`-stock-recommendations.xlsm**, where `netid` is your university-issued net identifier (i.e. the first part of your university-issued email address).

For a guided walkthrough of step-by-step instructions, see the project [checkpoints](checkpoints.md).

In addition to providing a good user experience, your submission should adhere to the following requirements, as detailed in the corresponding sections below:

  + Information Input Requirements
  + Validation Requirements
  + Information Output Requirements
  + Calculation Requirements

## Information Input Requirements

The system should prompt the user to input one or more stock symbols (e.g. `"MSFT"`, `"AAPL"`, etc.).

![an example user interface which prompts the user to input a stock symbol into cell E11 and then press a command button to initiate the recommendation process](example-interface.png)

The system should capture inputs via your choice of input mechanism, whether it be cell value(s), input boxes, or some other means. The system may optionally prompt the user to specify additional inputs such as risk tolerance and other trading preferences, as desired and applicable.

After entering desired inputs, the user should be able to click a command button to trigger the recommendation process.

## Validation Requirements

Before requesting data from the Internet, the system should first perform preliminary validations on user inputs. For example, it should ensure stock symbols are `String` datatypes and less than around six characters long.

If preliminary validations are not satisfied, the system should display a friendly error message like "Oh, expecting a properly-formed stock symbol like 'MSFT'. Please try again." and stop execution.

Otherwise, if preliminary validations are satisfied, the system should proceed to issue a GET request to the [AlphaVantage API](https://www.alphavantage.co/documentation/) to retrieve corresponding stock market data.

When the system makes an HTTP request for that stock symbol's trading data, if the stock symbol is not found or there is an error message returned by the API server, the system should display a friendly error message like "Sorry, couldn't find any trading data for that stock symbol", and it should stop program execution to allow the user to try again.

## Information Output Requirements

After receiving an API response, the system should write historical stock prices to one or more worksheet(s). If the system processes only a single stock symbol at a time, the system may use a single sheet named something like "Data". Whereas if the system processes multiple stock symbols at a time, for each stock symbol, the system should write historical trading data to a corresponding worksheet that is named after the stock symbol. If writing multiple sheets of data, the system should have a way of cleaning-up to prevent uncontrolled proliferation of new sheets.

After writing historical data to a spreadsheet, the system should perform calculations (see "Calculation Requirements" section below) to produce a recommendation as to whether or not the client should buy the stock, and optionally what quantity to purchase. The nature of the recommendation for each symbol can be binary (e.g. "Buy" or "No Buy"), qualitative (e.g. a "Low", "Medium", or "High" level of confidence), or quantitative (i.e. some numeric rating scale). The final recommendations can be displayed using your choice of output mechanism, whether it be cell values, message boxes, or some other means.

![a screenshot of a message box showing a recommendation for the user to buy the stock](example-recommendation.png)

Anywhere price-related information is displayed, it should be formatted as USD with a dollar sign and two decimal places. This includes on the "Data" sheet as well as in final recommendations.

## Calculation Requirements

You are free to develop your own custom recommendation algorithm. This is perhaps one of the most fun and creative parts of this project. :smiley:

One simple example algorithm would be (in pseudocode): If the stock's latest closing price is less than 20% above its 52-week low, "Buy", else "Don't Buy".

## Submission Instructions

Upload your workbook file to Canvas:

  + [Section 40 Project Upload](https://georgetown.instructure.com/courses/54379/assignments/123538)
  + [Section 41 Project Upload](https://georgetown.instructure.com/courses/54380/assignments/123537)

## Evaluation Methodology

Submissions will be evaluated based on ability to meet each of the component requirements (see corresponding sections above for detailed instructions):

Category | Weight
--- | ---
User Experience and Instructions | 20%
Information	Input Requirements | 15%
Validation Requirements (Performs Preliminary Validations) | 15%
Calculation Requirements (Issues API Request) | 8%
Validation Requirements (Handles API Response Errors) | 15%
Information	Output Requirements (Writes Data Sheet of Historical Prices) | 7%
Calculation Requirements (Appropriateness of Custom Algorithm) | 7%
Information	Output Requirements (Displays Final Recommendations) | 13%

This rubric is tentative, and may be subject to slight adjustments during the grading process.

The professor reserves the right to award extra credit in recognition of particularly effective user experiences. Common elements that may be eligible for extra credit include: simplicity of user interface, clarity of user instructions, and exceeding basic expectations set forth here for auto-updating charts and graphics. Also eligible for extra credit are submissions which expand the scope of this project to handle multiple stock symbols and compare across stock symbols.
