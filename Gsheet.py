import json
import re
import datetime
import gspread
from oauth2client.client import SignedJwtAssertionCredentials

import socket
import ib
from time import sleep

from ib.ext.Contract import Contract
from ib.opt import ibConnection, message
from time import sleep

#
#  Classes
#

class gsheet(object):

    def __init__(self):
        json_key = json.load(open(keyFile))
        scope = ['https://spreadsheets.google.com/feeds']
        credentials = SignedJwtAssertionCredentials(json_key['client_email'], json_key['private_key'].encode(), scope)
        self.gs = gspread.authorize(credentials)

    def wks(self,wbk,sheet):
        w = self.gs.open(wbk).worksheet(title=sheet)
        return w

class Downloader(object):

    def __init__(self,port):
        self.tws = ibConnection('localhost', port, 0)
        self.tws.registerAll(self.reply_handler)
        self.tws.connect()
        self._reqId = 1 # current request id

    def reply_handler(self,msg):
        print 'Reply:', msg
        if msg.typeName == 'accountSummary':
            print msg.tag
            #accountUpdate(msg)
        if msg.typeName == 'historicalData':
            if msg.close > 0:
                row = msg.reqId-100
                print row,msg.close
                #HL1.update_cell(row,24,msg.close)

    def requestData(self,contract):
        self.tws.reqMktData(self._reqId, contract, '', 1)
        self._reqId+=1

    def requestHistoricalData(self,contract,id):
        now = datetime.datetime.now().strftime('%Y%m%d %H:%M:%S')
        self.tws.reqHistoricalData(id, contract, now, '3600 S', '1 hour','MIDPOINT',0,1)

    def requestPositions(self):
        self.tws.reqPositions()

    def disconnect(self):
        self.tws.disconnect()

    def requestAccounts(self):
        self.tws.reqAccountSummary(self._reqId,'All','NetLiquidation')
        self._reqId+=1

def accountUpdate(msg):
    now = datetime.datetime.now().strftime('%Y-%m-%d')

    if msg.account == 'U1549891':
        if msg.tag == 'NetLiquidation':
            Dashboard.update_acell('b38', msg.value)
            Dashboard.update_acell('c38', now)
    if msg.account == 'U1553880':
        if msg.tag == 'NetLiquidation':
            Dashboard.update_acell('b40', msg.value)
            Dashboard.update_acell('c40', now)
    if msg.account == 'U1551005':
        if msg.tag == 'NetLiquidation':
            Dashboard.update_acell('b39', msg.value)
            Dashboard.update_acell('c39', now)
    if msg.account == 'U1552261':
        if msg.tag == 'NetLiquidation':
            Dashboard.update_acell('b41', msg.value)
            Dashboard.update_acell('c41', now)

def make_contract(symbol, sec_type="STK", exch="SMART", prim_exch="NYSE", curr="USD"):

    Contract.m_symbol = symbol
    Contract.m_secType = sec_type
    Contract.m_exchange = exch
    Contract.m_primaryExch = prim_exch
    Contract.m_currency = curr
    return Contract

if __name__=='__main__':

    #
    #  Assign Constants
    #

    keyFile = 'C:\Users\Roland\Google Drive\PyExchange-985ba7b9d627.json'

    #
    # Main program segment
    #

    gc = gsheet()
    global Dashboard
    global HL1
    Dashboard = gc.wks("2015b Investment workbook","Dashboard")
    HL1 = gc.wks("2015b Investment workbook","HL1 Log")

    val = Dashboard.acell('b38')
    print val

    val = HL1.acell("n5")
    print(val)

    dlR = Downloader(7496)
    dlN = Downloader(7497)

    dlR.requestPositions()
    dlR.requestAccounts()
    dlN.requestPositions()
    dlN.requestAccounts()

    p = HL1.col_values(25)
    print p
    for row in range(5,len(p)+1):
        if p[row-1] is not None:
            c = Contract()
            c.m_symbol = HL1.cell(row,2).value
            c.m_secType = 'OPT'
            c.m_exchange = 'SMART'
            c.m_currency = 'USD'
            c.m_expiry = HL1.cell(row,4).value
            c.m_strike = HL1.cell(row,5).value
            c.m_right = HL1.cell(row,6).value
            print c
            dlR.requestHistoricalData(c,100+row)


    sleep(10)

    dlR.disconnect()
    dlN.disconnect()

