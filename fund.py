import datetime
import threading

def func():
    #####################################################################################################################
    import requests
    import pandas as pd
    import numpy as np
    from bs4 import BeautifulSoup
    import datetime

    
    sheet=pd.read_csv("C:/Users/STEVEN/Desktop/基金淨值/FUND_WORK_SHEET.csv" , encoding="big5")
    sheet
    fund=[]
    outlook=[]
    for i in range(34):
        outlook.append(sheet.CODE[i])
    for i in range(34):
        fund.append(sheet.NAV[i])
    # display(outlook,fund)

    ####################################################################################################################
    import pandas as pd
    import datetime 
    from datetime import date
    timestamp=date.today()
    t=str(timestamp).replace("-","_")
    tt=t[5:].replace("_","-")
    z=str(datetime.datetime.now())
    FUND_NAME=[]
    isin_code=[]
    today=[]
    price=[]
    last_price=[]
    change=[]
    head = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.61 Safari/537.36'}
    ##淨值以及檢測日期
    for i in fund:
        data=[]
        url=i
        page = requests.get(url,headers = head)
        soup = BeautifulSoup(page.text, "html.parser")
        tag=soup.find_all("td" , align="right")
        for i in tag:
            data.append(i.text)    
        now=data[6]
        net_worth=data[7]
        yesterday_price=data[9]
        today.append(now)
        price.append(net_worth)
        last_price.append(yesterday_price)
        change.append(((float(net_worth)/float(yesterday_price)-1)*100))
    # #ISIN_CODE 以及基金名稱
    for j in outlook:
        name=[]
        info=[]
        web=j
        head = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.61 Safari/537.36'}
        page= requests.get(web,headers = head)
        soup_ = BeautifulSoup(page.text, "html.parser")
        raw_info=soup_.find_all("td",class_="FieldContent")
        fund_name=soup_.find_all("td",class_="componentTitle")
        for i in raw_info:
            info.append(i.text)
        code=info[4]
        isin_code.append(code)
        for j in fund_name:
            name.append(j.text)
        FUND_NAME.append(name)
    import pandas as pd
    import datetime 
    from datetime import date
    timestamp=date.today()
    t=str(timestamp).replace("-","_")
    import pandas as pd
    col = ['FUND_NAME','ISIN_CODE',"DATE","PRICE","LAST_DATE_PRICE","%CHAGE"]
    df1 = pd.DataFrame(list(zip(FUND_NAME,isin_code,today,price,last_price,change)), columns=col)
    # df.to_excel("C:/Users/STEVEN/Desktop/Fund_info_"+t+".xlsx" ,sheet_name="Fund", encoding="big5")

###################################################################################################################

    data=["https://www.moneydj.com/funddj/ya/yp010000.djhtm?a=ACML01",\
          "https://www.moneydj.com/funddj/ya/yp010000.djhtm?a=ACNC10",\
          "https://www.moneydj.com/funddj/yp/yp010000.djhtm?a=ACNC06",\
         "https://www.moneydj.com/funddj/yp/yp010000.djhtm?a=ACPS04",\
         "https://www.moneydj.com/funddj/ya/yp010000.djhtm?a=ACML04",\
         "https://www.moneydj.com/funddj/yp/yp010000.djhtm?a=ACDS13",\
         "https://www.moneydj.com/funddj/ya/yp010000.djhtm?a=ACCA02",\
         "https://www.moneydj.com/funddj/yp/yp010000.djhtm?a=ACDD02",\
         "https://www.moneydj.com/funddj/yp/yp010000.djhtm?a=ACFT16",\
         "https://www.moneydj.com/funddj/yp/yp010000.djhtm?a=ACDD19",\
         "https://www.moneydj.com/funddj/yp/yp010000.djhtm?a=ACCB20",\
         "https://www.moneydj.com/funddj/yp/yp010000.djhtm?a=ACTC02",\
         "https://www.moneydj.com/funddj/yp/yp010000.djhtm?a=ACYC03",\
         "https://www.moneydj.com/funddj/yp/yp010000.djhtm?a=ACFP39",\
         "https://www.moneydj.com/funddj/yp/yp010000.djhtm?a=ACDD26"]
    NAME=[]
    DATE=[]
    REV=[]
    LAST=[]
    LAST_DATE=[]

    for i in data:
        url=i
    
        page = requests.get(url,headers = head)
        soup = BeautifulSoup(page.text, "html.parser")
        name=soup.find_all("title")[0]
        date=soup.find_all("td", class_="t3n0c1")[0]
        rev=soup.find_all("td", class_="t3n1")[0]
        last_rev=soup.find_all("td", class_="t3n1_rev")[0]
        last_date=soup.find_all("td" ,class_="t3n0c1_rev")[0]
        for k in name:
            NAME.append(k.replace("-淨值表-基金-MoneyDJ理財網",""))
        for j in date:
            DATE.append(j)
        for i in rev:
            REV.append(i)
        for v in last_rev:
            LAST.append(v)
        for x in last_date:
            LAST_DATE.append(x)
        
#     display(DATE)
        
    import pandas as pd
    col = ['FUND_NAME',"DATE","PRICE","LAST_DATE","LAST_PRICE"]
    df2 = pd.DataFrame(list(zip(NAME,DATE,REV,LAST_DATE,LAST)),columns=col)
    # df_money.to_excel("C:/Users/STEVEN/Desktop/Fund_info_"+t+".xlsx",sheet_name="MoneyDJ" , encoding="big5")




    writer = pd.ExcelWriter("C:/Users/STEVEN/Desktop/基金淨值/Fund_info_"+t+".xlsx")
    df1.to_excel(writer,'fund')
    df2.to_excel(writer,'moneydj')
    writer.save()




#####################################################################################################################################
#####################################################################################################################################
#####################################################################################################################################
