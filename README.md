# VBA-GetJSON-RESTApi
Sample Excel VBA code to get JSON response from a JSON REST API
This example uses an API to return share values from https://www.alphavantage.co

In order to run the code you will need to register for an API key at https://www.alphavantage.co and assign to the VBA code const API_SECRET_KEY

This code is provided as is without warranty of any kind and use is at your own risk.

Source files are provided in TXT format which can be pasted into a new module in the excel VBA editor

The following steps are based on Excel 2016
goto url https://www.alphavantage.co and register for API key

Access the VBA editor
Add dependencies via \Tools References
Microsoft ActiveX Data Objects 2.8 library
Microsoft Scription Runtime

Macro name 
Right click modules
Insert New Module and name AlphaAdvantage
paste code from alphaAdvantage.txt
Insert New Module and name VBAJson
paste code from VbaJson.txt

open module AlphaAdvantage
at top of module API_SECRET_KEY add your API key from above

At the bottom of module AlphaAdvantage is a sub Test.
place cursor within and run code

View Immediate window should show Json data recieved from API
amend code as required for your symbols, assign data to spreadsheet cells etc.
