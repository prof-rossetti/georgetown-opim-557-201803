# Software APIs Overview

> Before you read this document, please first read about [Computer Networks](computer-networks.md).

Humans often interface with software manually, using a **Graphical User Interface (GUI)** which most likely includes buttons, navigation menus, drag-and-drop functionality, etc.

However, many applications also allow programmatic use. By specifying an **Application Programming Interface (API)**, or instructions on how to use the software programmatically, applications allow both humans and other applications to send data to the application and receive data from it.

It is not uncommon for a system to also use its own public API to perform its own desired functionality.

Most of todays popular APIs are **Web Services** which accept HTTP requests at specified URLs and return responses to fulfill those requests. Here are some example APIs and API providers:

 + [New York Times APIs](http://developer.nytimes.com/docs)
 + [Google APIs](https://developers.google.com/apis-explorer/#p/)
 + [Twitter APIs](https://dev.twitter.com/rest/public)
 + [Facebook Social Graph API](https://developers.facebook.com/docs/graph-api)
 + [Instagram API](https://instagram.com/developer/endpoints/)
 + [Foursquare API](https://developer.foursquare.com/docs/)
 + [GitHub API](https://developer.github.com/v3/)
 + [Yelp API](https://www.yelp.com/developers/documentation/v2/overview)
 + [Flickr API](https://www.flickr.com/services/api/)
 + [Getty Images API](http://developers.gettyimages.com/en/)
 + [US Federal Elections Commission API](https://api.open.fec.gov/developers)
 + [Alpha Vantage (Stock Market) API](https://www.alphavantage.co/documentation/)

### Authentication

Many web services require developers to first register to obtain valid credentials in the form of an **API Key** (i.e. a secret token string) and subsequently authenticate by providing the key alongside each API request.

### Response Formats

The most common format for API response data is JSON, but some APIs alternatively or additionally provide response data in XML or CSV format.

Example CSV:

```csv
city,name,league
New York,Mets,Major
New York,Yankees,Major
Boston,Red Sox,Major
Washington,Nationals,Major
New Haven,Ravens,Minor
```

Example JSON:

```js
[
  {"city": "New York", "name": "Yankees", "league":"Major"},
  {"city": "New York", "name": "Mets", "league":"Major"},
  {"city": "Boston", "name": "Red Sox", "league":"Major"},
  {"city": "Washington", "name": "Nationals", "league":"Major"},
  {"city": "New Haven", "name": "Ravens", "league":"Minor"}
]
```

Example XML:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<teams>
  <team>
    <city>New York</city>
    <league>Major</league>
    <name>Yankees</name>
  </team>
  <team>
    <city>New York</city>
    <league>Major</league>
    <name>Mets</name>
  </team>
  <team>
    <city>Boston</city>
    <league>Major</league>
    <name>Red Sox</name>
  </team>
  <team>
    <city>Washington</city>
    <league>Major</league>
    <name>Nationals</name>
  </team>
  <team>
    <city>New Haven</city>
    <league>Minor</league>
    <name>Ravens</name>
  </team>
</teams>
```

### Request Parameters

Many APIs allow you to specify URL parameters along with your HTTP request. These URL parameters are appended to the end of the API's base URL, starting with a single question mark (`?`) to denote the rest of the URL contains parameters. Then each parameter follows a convention where the name of the parameter is followed by an equal sign (`=`), which is followed by the desired parameter value. If there are multiple parameters, subsequent parameters after the first are separated by the ampersand character `&`.

Example request URL: https://www.alphavantage.co/query?function=TIME_SERIES_DAILY&symbol=MSFT&outputsize=compact&apikey=demo.

In this example, `https://www.alphavantage.co/query` is the base URL. And `function`, `symbol`, `outputsize`, and `apikey` are the names of URL parameters.
