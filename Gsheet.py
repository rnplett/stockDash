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
                row = msg.reqId-100-5
                p[row].value = msg.close
                print row, msg.close, p[row].value
                #HL1.update_cell(row,24,msg.close)
        if msg.typeName == 'tickOptionComputation':
            row = msg.tickerId-100-5
            p[row].value = msg.optPrice
            d[row].value = msg.delta
            print row, p[row].value, d[row].value
        if msg.typeName == 'tickPrice':
            row = msg.tickerId-100-5
            p[row].value = msg.price
            print row, p[row].value, d[row].value

    def requestData(self,contract,id):
        self.tws.reqMktData(id, contract, '', 1)
        #self._reqId+=1

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

def make_contract(symbol, sec_type="STK", exch="SMART", curr="USD", expiry="", strike="", right=""):

    Contract.m_symbol = symbol
    Contract.m_secType = sec_type
    Contract.m_exchange = exch
    Contract.m_currency = curr
    if sec_type == 'OPT':
        Contract.m_expiry = expiry
        Contract.m_strike = strike
        Contract.m_right = right

    return Contract

if __name__=='__main__':

    #
    #  Assign Constants
    #

    keyFile = 'C:\Users\Roland\Google Drive\PyExchange-985ba7b9d627.json'
    workbook = '2015b Investment workbook'

    logWks = 'HL1 Log'
    symbolCol = 2
    expiryCol = 4
    strikeCol = 5
    rightCol = 6
    priceCol = 25
    deltaCol = 30
    headerRow = 5

    #
    # Main program segment
    #

    gc = gsheet()

    #  Dashboard Update
    #
    Dashboard = gc.wks(workbook,"Dashboard")

    val = Dashboard.acell('b38')
    print val

    dlR = Downloader(7496)
    dlN = Downloader(7497)

    #dlR.requestPositions()
    #dlR.requestAccounts()
    #dlN.requestPositions()
    #dlN.requestAccounts()

    #  Log Updates
    #
    HL1 = gc.wks(workbook,logWks)
    p = HL1.range(HL1.get_addr_int(headerRow,priceCol)+":"+HL1.get_addr_int(1000,priceCol))
    d = HL1.range(HL1.get_addr_int(headerRow,deltaCol)+":"+HL1.get_addr_int(1000,deltaCol))
    print p
    for cl in p:
        if cl.value is not '':
            c = make_contract(HL1.cell(cl.row,symbolCol).value,'OPT','SMART','USD', \
                HL1.cell(cl.row,expiryCol).value,HL1.cell(cl.row,strikeCol).value,HL1.cell(cl.row,rightCol).value)
            #dlR.requestHistoricalData(c,100+cl.row)
            dlR.requestData(c,100+cl.row)

    sleep(10)

    print p
    HL1.update_cells(p)
    HL1.update_cells(d)

    dlR.disconnect()
    dlN.disconnect()

