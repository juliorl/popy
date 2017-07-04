# popy
#### Simple fixed income portfolio tools in Python
___

####Functions

*read_portfolio (excelFile, excelSheet)*

Returns a 'portfolio' pandas dataframe.
Load an Excel sheet with securities. Takes 'excelFile' and 'excelSheet' with columns:
issueCode, couponRate, maturityDate, tradeDate, lastCoupon, couponPeriod, faceValue,
YTM, currency, account. Order is important, starting at R0C0.

*read_security (x)*

Returns a 'security' dict. Reads security in position 'x' from "portfolio" dataframe.

*generate_cashflows (lastCoupon, maturityDate, couponPeriod, couponRate, faceValue)*

Returns a 'cashflows' pandas dataframe with dates and payments.
ONLY FIXED RATE BONDS supported, at the moment.

*write_results (outputFile)*

Writes cashflow of a security in 'outputFile' Excel file. Also converts QuantLib dates to timestamp, and removes hours and keeps sonly dates.
