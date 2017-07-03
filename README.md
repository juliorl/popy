# popy
#### Simple fixed income portfolio tools in Python
___
##### Version 0.0.1
TL;DR: Initial commit. Basic functions.

**Functions**

*read_portfolio(excelFile, excelSheet)*

Load an Excel sheet with securities.
Takes 'excelFile' and 'excelSheet' with the following columns:
issueCode, couponRate, maturityDate, tradeDate, lastCoupon, couponPeriod,
faceValue, YTM, currency, account. Order is important, starting at R0C0.
Returns a 'portfolio' pandas dataframe.

*read_security(x)*

Read 'x' line from "portfolio" dataframe and set up variables.
Returns a 'security' dict with security x's info.

*generate_cashflows(lastCoupon, maturityDate, couponPeriod, couponRate, faceValue)*

Generates cashflows and, of course, dates of said cashflows for securities.
ONLY FIXED RATE BONDS supported, at the moment.
Returns a 'cashflows' pandas dataframe with dates and payments.
