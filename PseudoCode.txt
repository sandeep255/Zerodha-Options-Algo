import logging
import xlrd
from kiteconnect import KiteConnect
from optionchain_stream import OptionChain #for getting the option chain

class FTS():
    #initialization based on tradesheet(we will get the zerodha entry from this file) and credentials to connect
    def __init__(self, tradeSheetFile = "tradesheet.xlsx", credentialsFile = "credentials.xlsx"):#This function will auto-execute 
        self.tradeSheetFile = tradeSheetFile
        self.credentialsFile = credentialsFile
        #session login
        print("Logger executed")
        logging.basicConfig(level=logging.DEBUG)
        print("Kite connect executing, using your api_key") #Print statements for debugging
        kite = KiteConnect(api_key="your_api_key")
        print("Session generating, using request token and secret")
        data = kite.generate_session("request_token_here", api_secret="your_secret")
        kite.set_access_token(data["access_token"])

    def readData: #using this function for reading the given data and storing it in array format
        print("Reading data from excel")
        book = xlrd.open_workbook('Zerodha Entry.xlsx')
        sheet = book.sheet_by_name('Sheet1')
        data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
        for i in range(1, sheet.nrows): #We are going through the excel file and getting values from the data provided by the user
            instrument = data[i][0] #In this manner the user need not code and just provide us with the script name
            premium = data[i][1] # the premium at which he/she wants to purchase it.
            distance = data[i][2] #The distance for hedging, by taking distance of 7 or more we can ensure that the delta of the hedge is half of the strike_price
            buy_qtr = data[i][3] #The quantity we want to buy
            sell_qtr = data[i][4] # The quantity we want to sell, buy and sell quantity is imp as by this combination alone we can setup debit spread or credit spread
            if premium==0:#This condition was specified that if the premium is 0 do not buy or sell
                continue;
            trade(instrument)

    def getStrikePrice(self,instrument, premium):
        #We will get the option chain data from Zerodha
        print("Getting option chain data")
        #Below code ensures the option data fetch
        OptionStream = OptionChain(instrument, "option_expiry_date in yyyy-mm-dd format", "api_key",
                    "api_secret=None", "request_token=None", "access_token=None", underlying=False) #session has to be validated again here
        OptionStream.sync_instruments() #This needs to be done only once a day

        #For selecting strike_price we will traverse through the JSON response of the OptionStream
        while(ltp>premium):
            # We will move to the next strike_price
            # if ltp < or = premium:
            #    strike_price = we will fetch the strike_price correspunding to the ltp of strike_price where premium is > or = ltp
            #    premium = ltp of strike_price
        return strike_price

    def getHedgeStrikePrice(self,instrument, distance, strike_price):
        if instrument is Nifty:
            hedge_strike_price = strike_price + distance*50 #distance multiplied by the lot size
        if instrument is Bank nifty:
            hedge_strike_price = strike_price + distance*100
        return hedge_strike_price #I had one concern regarding this distance method, it can be illiquid
        #but this issue would not arise if we are trading in index options and that too in weekly expiry.

    def placeOrder(self,instrument, strike_price, hedge_strike_price, buy_qtr, sell_qtr):
        #first buy the hedge_strike_price then the strike_price
        try:
            order_id = kite.place_order(variety=kite.VARIETY_REGULAR,
            tradingsymbol=instrument+strike_price+CE/PE, #CE OR PE Value can be taken in from the intrument name provided by the user
            exchange=kite.EXCHANGE_NSE, #I am not aware of the trading symbol in which kite api accepts it that is why I have left it as CE/PE
            transaction_type=kite.TRANSACTION_TYPE_BUY,
            quantity=buy_qtr,
            order_type=kite.ORDER_TYPE_limit,
            price = premium, #strike_price has been selected in getStrikePrice and the premium is updated to the ltp of the selected strike_price
            product=kite.PRODUCT_CNC,
            validity=kite.VALIDITY_DAY)

            logging.info("Order placed. ID is: {}".format(order_id))
        except Exception as e:
            logging.info("Order placement failed: {}".format(e.message))

        #buying the strike_price
        try:
            order_id = kite.place_order(variety=kite.VARIETY_REGULAR,
            tradingsymbol=instrument+hedge_strike_price+CE/PE,
            exchange=kite.EXCHANGE_NSE,
            transaction_type=kite.TRANSACTION_TYPE_SELL,
            quantity=sell_qtr,
            order_type=kite.ORDER_TYPE_limit,
            price = premium, #strike_price has been selected in getStrikePrice and the premium is updated to the ltp of the selected strike_price
            product=kite.PRODUCT_CNC,
            validity=kite.VALIDITY_DAY)

            logging.info("Order placed. ID is: {}".format(order_id))
        except Exception as e:
            logging.info("Order placement failed: {}".format(e.message))

    def trade(self,instrument):
        print("Getting strike price of",instrument)
        strike_price = self.getStrikePrice(instrument, premium) #Function is called for strike price selection

        print("Getting hegde strike price of",instrument)
        hedge_strike_price = self.getHedgeStrikePrice(instrument, distance, strike_price) #function is called for hedge_strike_price selection

        print("Placing order of", instrument,"at", strike_price,"with buying quantity",buy_qtr,"selling quantity",sell_qtr,"hedging strike price is",hedge_strike_price)
        placeOrder(instrument, strike_price, hedge_strike_price, buy_qtr, sell_qtr)#This is the final step where all the values have been found out and the order is placed


#Calling the class
obj = FTS()
obj.readData() #Calling read data function
