import b as binance
import time
import xlwings as xw
from playsound import playsound

binance.set(open("key.txt", "r").read())

def add_to_excel(sheet, money, totalAssetValue, coinInPortfolio, costOfTrade, buy_sell, counter):
    import time
    sheet.range(f'A{counter}').value = time.strftime("%d/%m/%Y")
    sheet.range(f'B{counter}').value = time.strftime("%H:%M:%S")
    sheet.range(f'C{counter}').value = money
    sheet.range(f'D{counter}').value = totalAssetValue
    sheet.range(f'E{counter}').value = coinInPortfolio
    sheet.range(f'F{counter}').value = costOfTrade
    sheet.range(f'G{counter}').value = buy_sell



xlsht = xw.Book('btc-bot.xlsx').sheets[0]
coin = "BTCUSDT"
threshold = float(input("Enter the threshold value: "))
total_money = float(input("Enter the total money to invest: "))
n = float(input("Enter the amount of coin to buy on being triggered : "))
add_to_excel(xlsht, "money", "totalAssetValue", "coinInPortfolio", "CostAtTrade", "buy/sell", 1)
# priceList = binance.prices()
price = float(binance.prices()[f'{coin}'].strip())
last_traded_price = price
print("price : ", price)
cuo = 0
xlcntr = 2
timerCntr = 0
xlsht.range('I2').value = "PriceOfOneCoin"
xlsht.range('I3').value = "coins in portfolio"
xlsht.range('I4').value = "last_traded_price"
xlsht.range('I5').value = "threshold - DeltaTrigger"
xlsht.range('I6').value = "coins per transaction"
xlsht.range('I7').value = "xlcntr"
xlsht.range('I8').value = "timerCntr"
xlsht.range('I9').value = "date"
xlsht.range('I10').value = "time"
xlsht.range('J6').value = n
xlsht.range('J5').value = threshold
while True:
    timerCntr+=1
    try:
        price = float(binance.prices()[f'{coin}'].strip())
    except Exception as e:
        print(e)
        time.sleep(5)
        continue
    if price <= last_traded_price - threshold and total_money >= n*price:
        # buy
        cuo += n
        total_money -= n*price
        last_traded_price = price
        add_to_excel(xlsht, total_money, total_money+(cuo*price), cuo, price, "buy", xlcntr)
        add_to_excel(xlsht, cuo, cuo*price, coin, price, "buy", xlcntr)
        xlcntr += 1
        playsound('Sosumi.aiff')

    elif price >= last_traded_price + threshold and cuo > n:
        # sell
        cuo -= n
        last_traded_price = price
        total_money += last_traded_price*n
        add_to_excel(xlsht, total_money, total_money+(cuo*price), cuo, last_traded_price, "sell", xlcntr)
        add_to_excel(xlsht, cuo, cuo*price, coin, price, "sell", xlcntr)
        xlcntr += 1
        playsound('Sosumi.aiff')

    else:
        pass

    if timerCntr%5 == 0:
        print("price : ", price)
        print("cuo : ", cuo)
        print("last_traded_price : ", last_traded_price)
        print("threshold : ", threshold)
        print("n : ", n)
        print("xlcntr : ", xlcntr)
        print("timerCntr : ", timerCntr)
        print("\n")
        timerCntr = 0
    
    else:
        pass

    if xlsht.range('J4').value == 0:
        last_traded_price = price

        
    n = xlsht.range(f'J6').value
    threshold = xlsht.range(f'J5').value
    xlsht.range('J2').value = price
    xlsht.range('J3').value = cuo
    xlsht.range('J4').value = last_traded_price
    xlsht.range('J7').value = xlcntr
    xlsht.range('J8').value = timerCntr
    xlsht.range('J9').value = time.strftime("%d/%m/%Y")
    xlsht.range('J10').value = time.strftime("%H:%M:%S")
    time.sleep(1)