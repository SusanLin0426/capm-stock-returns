#各產業股票CAPM模型分析
#Group 13 



#input : 各產業股票報酬率、市場報酬率
#output : 各產業回歸模型的係數α 值(額外報酬)、β值(系統性風險係數) 

#1.取得資料
#2.資料轉檔
#3.資料整理
#4.跑回歸，輸出結果



#import module
from selenium import webdriver                                                #自動化開啟瀏覽器、點擊按鈕、取得網站內容
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd                                                           #資料處理及資料分析工具
import csv                                                                    #csv套件可以進行csv格式檔案的儲存或處理(此程式中用來讀取csv檔)          
import numpy                                                                  #矩陣和陣列運算
from sklearn.linear_model import LinearRegression                             #跑回歸的套件
import matplotlib.pyplot as plt                                               #畫圖                    
from matplotlib.font_manager import FontProperties                            #管理字型的套件(需要指定中文字型)




#到櫃買中心下載資料(市場報酬率、各產業股票報酬率)

#櫃買中心網址
url = "https://www.tpex.org.tw/web/stock/statistics/monthly/monthly_rpt_mkt_info_10.php?l=zh-tw"

#設定點擊:
options = Options()
options.add_argument("--disable-notifications")    # 取消所有的alert彈出視窗(引用chrome模組底下的Options類別，設定不啟用通知，避免跳出訊息視窗，阻礙之後的自動化執行)
browser = webdriver.Chrome( ChromeDriverManager().install(), chrome_options=options)  #指定瀏覽器為Chrome
browser.get(url)                                                               #前往指定網站  (自動開啟Chrome並前往櫃買中心網站)
button = browser.find_element_by_class_name("table-text-over")                 #在selenium套件中用find_element_by_class_name方法利用class name指定要按的按鍵
#click 
button.click()                                                                 #用click方法點擊按鈕
 


#-------------------------------------------------------------------------------------------


# 資料轉檔 (xlxs轉csv)


#define excel轉csv檔的function
def xlxs_to_csv_pd(xls_file):
    data_xls = pd.read_excel(xls_file, index_col=0)                             #利用pandas套件讀檔
    csv_file = xls_file.split('.')[0]                                           #把excel檔檔名用"."作為分割符號進行分割，只取出第0項
    #print(xls_file)  => C:\Users\user\Downloads\投資報酬率11005.xls
    #print(csv_file)  => C:\Users\user\Downloads\投資報酬率11005
    data_xls.to_csv(csv_file + '.csv', encoding='utf-8')                        #檔案: C:\\Users\\user\\Downloads\\投資報酬率11005.csv ， 編碼類型為utf-8


#使用function進行轉檔
xlxs_to_csv_pd("/Users/user/Downloads/投資報酬率11005.xls")                      #下載好的Excel在Downloads資料夾中



#----------------------------------------------------------------------------------------------


#資料整理

df = pd.read_csv('/Users/user/Downloads/投資報酬率11005.csv',encoding="utf-8")    #轉檔後的csv檔在Downloads資料夾中
x = df.iloc[3:24]                                                                #只要資料的第3~24列

data_df = pd.DataFrame(x)

data_df.to_csv("data_ok.csv")                                                    #把資料存成csv檔


#確認資料，print出產業別
with open('data_ok.csv','r',encoding="utf-8") as csvfile:
    reader = csv.reader(csvfile)
    column = [row[1] for row in reader]
print(column[3:24])



#-------------------------------------------------------------------------------------------------



#跑回歸

#取市場報酬資料
with open('data_ok.csv','r',encoding="utf-8") as csvfile:
    reader = csv.reader(csvfile)
    for i,rows in enumerate(reader):
        if i == 1:                                                                #讀取市場報酬所在的那一列
            row_mkt = rows
    mkt_data = [row_mkt[3],row_mkt[4],row_mkt[5]]                                 #根據數據位置取數值
   





# Y = inter + coef * X
# Ri =  α  +  β * Rm 
# Y:各產業股票報酬率，X:市場報酬率
Beta = []       #放斜率資料
Alpha = []      #放截距資料
stock =[]


#取股票報酬資料
with open('data_ok.csv','r',encoding="utf-8") as csvfile:
    reader = csv.reader(csvfile)
    for index,rows in enumerate(reader):
        for x in range(3,23):
            if index == x:
                row_stock = rows
                industry =  row_stock[1]                                          #每列的第一筆資料為產業別名稱
                stock_data = row_stock[3],row_stock[4],row_stock[5]               #每列的第3~5筆資料為該產業不同時間點的股票報酬率
                
                #跑回歸
                x = numpy.array(mkt_data)
                y = numpy.array(stock_data)
                stock_datas = " ".join(stock_data)                                #把tuple中元素用空格隔開
                stock+=stock_datas.split()
                lm = LinearRegression()
                lm.fit(numpy.reshape(mkt_data, (len(x), 1)), numpy.reshape(y, (len(y), 1)))
                print(lm.coef_)
                Beta.append(lm.coef_)                                              #每次做完一個產業的回歸把斜率值放到Beta這個list中
                print(lm.intercept_ )
                Alpha.append(lm.intercept_)                                        #每次做完一個產業的回歸把截距值放到Alpha這個list中
                print(stock_data)
                print(industry)




#CAPM模型的斜率項=beta(系統性風險(總體環境造成的市場風險)係數)，β 值越高代表該股票或投資組合波動相對總體市場的波動更加敏感，如果大於 1 代表該股票或投資組合的波動幅度較市場大，小於 1 代表波動幅度比市場小，等於 1 代表波動幅度與市場相同。
# 截距項=alpha = 資產的額外報酬，α 值是β 無法衡量到的超額報酬，如果實際報酬高於預期報酬，α 值則大於 0，代表組合表現的比預期好，反之則小於 0
#print(coef)   #斜率(beta)
#print(inter)  #截距項(alpha)







#plotting result:
#設定散佈圖標示文字(從csv資料檔中找類別名稱，存放到namelist中)
with open('data_ok.csv','r',encoding="utf-8") as csvfile:
    reader = csv.reader(csvfile)
    column = [row[1] for row in reader]                                               #各產業名稱位在第一行的第4到第22列
namelist = column[3:22]                       




#畫圖(有標示產業類別)
fig, ax = plt.subplots()
ax.scatter(Beta, Alpha,color='black')                                                 #把散佈圖中的點設定為黑色
ChineseFont = FontProperties(fname = 'C:\\Windows\\Fonts\\mingliu.ttc')               #指定字型，使用Windows內建的細明體
plt.title('各類股β、α 結果',FontProperties = ChineseFont)                              #設定散佈圖標題，指定字型為細明體
plt.xlabel('β值',FontProperties = ChineseFont)                                        #設定X軸座標名稱，指定字型為細明體
plt.ylabel('α值',FontProperties = ChineseFont)                                        #設定Y軸座標名稱，指定字型為細明體      
for i, txt in enumerate(namelist):
    ChineseFont = FontProperties(fname = 'C:\\Windows\\Fonts\\mingliu.ttc')           #用字型絕對路徑指定字型(Windows內建的細明體)
    ax.annotate(txt, (Beta[i],Alpha[i]), fontproperties = ChineseFont,color='blue')   #指定點旁邊的產業類別標示，設定文字顏色為藍色


#顯示圖象
plt.show()