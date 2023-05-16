"""
popy - simple fixed income portfolio tools in Python
https://github.com/sqrj
"""

import pandas as pd
import QuantLib as ql

def read_portfolio(excelFile, excelSheet):
    """ Returns a 'portfolio' pandas dataframe.
    Load an Excel sheet with securities. Takes 'excelFile' and 'excelSheet' with columns:
    issueCode, couponRate, maturityDate, tradeDate, lastCoupon, couponPeriod,
    faceValue, YTM, currency, account. Order is important, starting at R0C0.
    """
    portfolio = pd.read_excel(excelFile,excelSheet,dtype={'Issue':'str','Face value':'int'})
    return portfolio

def read_security(x):
    """ Returns a 'security' dict.
    Reads security in position 'x' from "portfolio" dataframe.
    """
    # read line from excelSheet in excelFile
    issueCode = portfolio['issueCode'][x]
    read_couponRate = portfolio['couponRate'][x]
    read_maturityDate = portfolio['maturityDate'][x]
    read_tradeDate = portfolio['tradeDate'][x]
    read_lastCoupon = portfolio['lastCoupon'][x]
    read_couponPeriod = portfolio['couponPeriod'][x]
    read_faceValue = portfolio['faceValue'][x]
    read_YTM = portfolio['YTM'][x]
    currency = portfolio['currency'][x]
    account = portfolio['account'][x]

    # transform loaded data, when needed
    couponRate = float(read_couponRate)
    maturityDate = ql.Date(read_maturityDate.day,read_maturityDate.month,read_maturityDate.year)
    tradeDate = ql.Date(read_tradeDate.day,read_tradeDate.month,read_tradeDate.year)
    ##placeholder for lastCoupon, for now assume = tradeDate
    lastCoupon = tradeDate
    couponPeriod = ql.Period(int(read_couponPeriod), ql.Days)
    faceValue = int(read_faceValue)
    YTM = float(read_YTM)

    # build dict with security info
    security = {
        'issueCode' : issueCode,
        'couponRate' : couponRate,
        'maturityDate' : maturityDate,
        'tradeDate' : tradeDate,
        'lastCoupon' : lastCoupon,
        'couponPeriod' : couponPeriod,
        'YTM' : YTM,
        'faceValue' : faceValue,
        'currency' : currency,
        'account' : account
    }

    return security # returns a dict

def generate_cashflows(lastCoupon, maturityDate, couponPeriod, couponRate, faceValue):
    """ Returns a 'cashflows' pandas dataframe with dates and payments.
        ONLY FIXED RATE BONDS supported, at the moment.
    """
    # these are constants used for building a QL schedule and a QL fixedRateBond object
    CALENDAR = ql.Argentina()
    DAY_COUNT = ql.ActualActual()
    BUSINESS_CONVENTION = ql.Following
    DATE_GENERATION = ql.DateGeneration.Forward
    MONTH_END = False
    SETTLEMENT_DAYS = 0
    coupon = [couponRate]

    # Generate schedule of coupons
    schedule = ql.Schedule(lastCoupon,maturityDate,couponPeriod,CALENDAR,BUSINESS_CONVENTION,BUSINESS_CONVENTION,DATE_GENERATION,MONTH_END)

    # Generate coupon payments as per schedule
    fixedRateBond = ql.FixedRateBond(SETTLEMENT_DAYS,faceValue,schedule,coupon,DAY_COUNT)

    # Create Pandas dataframe 'thisBond' with cashflows
    theseDates = list(schedule) # convert QuantLib schedule object to a list
    del(theseDates[0]) # delete issueDate that ql.Schedule adds first
    theseDates.append(theseDates[-1]) # duplicate the last element of the theseDates list for principal payment
    theseAmounts = {} # create empty dictionary
    for flow in fixedRateBond.cashflows():
        theseAmounts[flow.amount] = flow.amount()  # this will generate a dictionary, with weird keys!
    theseAmounts = list(theseAmounts.values()) # convert theseAmounts from dict to list
    cashflows = pd.DataFrame({'dates' : theseDates, 'payments' : theseAmounts}) # create dataframe with lists
    cashflows.payments = cashflows.payments.astype(int) # change payments column into integer

    return cashflows # this returns a Pandas dataframe

    def write_results (outputFile):
    """ Writes cashflow of a security in 'outputFile' Excel file.
    Also converts QuantLib dates to timestamp, and removes hours and keeps sonly dates.
    """
    outputSheet = security['issueCode']
    securityCashflow = security['cashflow'] # unwrap the cashflow dataframe from dict
    # convert QuantLib date objects in dataframe to pandas time stamps
    for i in securityCashflow.index:
        securityCashflow['dates'][i] = pd.to_datetime(str(securityCashflow['dates'][i]))
    # keep only dates (drop time) from pandas dataframe field
    securityCashflow['dates'] = securityCashflow['dates'].dt.date
    writer = pd.ExcelWriter(outputFile)
    securityCashflow.to_excel(writer, outputSheet, index = False)
    writer.save()
