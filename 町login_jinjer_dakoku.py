from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from tkinter import messagebox
import os
import sys
import tkinter as tk
import yaml
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from webdriver_manager.microsoft import EdgeChromiumDriverManager

root = None
resCloseUpDriver = None
resAskQuestion = None
resAnsRunningAppSearch = None

# #ブラウザChrome
# browser = 'executable_path="chromedriver.exe"'
# driver = webdriver.Chrome(browser)

#ブラウザEdgeを指定
# browser = 'executable_path="msedgedriver.exe"'
# driver = webdriver.Edge(browser)

#ブラウザEdge
#driver = webdriver.Edge(executable_path="msedgedriver.exe")

with open(os.path.dirname(sys.argv[0]) + '\config.yaml', 'r') as file:
    config = yaml.load(file, Loader=yaml.SafeLoader)
    CODE = config['code']
    MAIL = config['mail']
    PASS = config['pass']

def generateWindow():
    if __name__ == "__main__":
        # Windowを生成する。
        root = tk.Tk()
        #tk Windowの非表示
        root.withdraw()

def closeUpDriver(Botton):
    #jinjer閉じる確認
    resCloseUpDriver = messagebox.askquestion("アプリ終了確認", "「" + Botton + "」ボタンが既に押された状態でした。\n \
        jinjerを終了しますか? \n \n \
    「はい」: jinjerを終了します。 \n \
    「いいえ」 : jinjer開いたままにします。")
    generateWindow()
    return resCloseUpDriver

def getAskQuestion():
    #退勤打刻確認メッセージボックス表示 開いているアプリがあれば強制的に閉じます
    # resAskQuestion = messagebox.askquestion("jinjer退勤打刻確認", "退勤打刻しますか? \n \n \
    resAskQuestion = messagebox.askyesnocancel("jinjer退勤打刻確認", "退勤打刻しますか? \n \n \
    「はい」: 打刻後、シャットダウン。 \n \
    「いいえ」 :打刻せずにシャットダウン。 \n \
    「キャンセル」 :打刻、シャットダウン処理しません。")
    generateWindow()
    return resAskQuestion

def loginInfo():
    global driver
    #ブラウザEdge(最新インストール)
    driver = webdriver.Edge(EdgeChromiumDriverManager().install())

    url = "https://kintai.jinjer.biz/sign_in"
    driver.get(url)
    
    loginCode = driver.find_element(By.NAME,'company_code')
    loginEmail = driver.find_element(By.NAME,'email')
    loginPass = driver.find_element(By.NAME,'password')
    
    loginCode.send_keys(CODE)
    loginEmail.send_keys(MAIL)
    loginPass.send_keys(PASS)

    btn = driver.find_element(By.NAME,'button')

    btn.click()

try:
    #出勤退勤フラグ(引数が指定されている場合のみ)
    taikinFlg = '0'
    if len(sys.argv) > 1:
        taikinFlg = sys.argv[1]

    #スタートアップからの起動時（出勤処理）
    if  (taikinFlg == '1'):
        #ログイン処理　開始
        loginInfo()

        #出勤ボタンが表示されるまで待機
        wait = WebDriverWait(driver, 600)
        selectorS = '//*[@id="container"]/section/main/div[2]/div[2]/ul/li[1]/button'
        element = wait.until(EC.visibility_of_element_located((By.XPATH, selectorS)))

        #出勤ボタンの要素を取得
        elemBtn = driver.find_element(By.XPATH, selectorS)

        #出勤ボタンが押せる場合、出勤打刻し、ブラウザを閉じる。
        if (elemBtn.is_enabled() == True):
            #出勤ボタン押下
            element.click()
            
            #退勤ボタンが押せる状態になったらブラウザを終了
            #退勤ボタンが表示されるまで待機
            wait = WebDriverWait(driver, 600)
            selectorE = '/html/body/div[1]/div[4]/div[1]/section/main/div[2]/div[2]/ul/li[2]/button'
            element = wait.until(EC.visibility_of_element_located((By.XPATH, selectorE)))
            #退勤ボタンの要素を取得
            elemBtn = driver.find_element(By.XPATH, selectorE)
            #退勤ボタンが押せる場合
            if (elemBtn.is_enabled()) == True :
                #ブラウザを終了する。
                driver.quit()
        else:
            #出勤ボタンが押せ無い状態（すでに「出勤」ボタンが押された状態。）
            #jinjar確認するか（する「はい」→jinjar開いたまま、しない「いいえ」→jinjar閉じる）
            ansCloseUpDriver = closeUpDriver('出勤')
            if ansCloseUpDriver == 'yes':
                #ブラウザを終了する。
                driver.quit()
            else:
                pass
    else:
        #退勤処理
        ansAskQuestion = getAskQuestion()
        if ansAskQuestion == True:
            #print('はい選択。退勤打刻します。打刻後、シャットダウンします。')
            #ログイン処理　開始
            loginInfo()
            #退勤ボタンが表示されるまで待機
            wait = WebDriverWait(driver, 600)
            selectorE = '/html/body/div[1]/div[4]/div[1]/section/main/div[2]/div[2]/ul/li[2]/button'
            element = wait.until(EC.visibility_of_element_located((By.XPATH, selectorE)))

            #退勤ボタンの要素を取得
            elemBtn = driver.find_element(By.XPATH, selectorE)

            # 退勤ボタンが押せる場合
            if (elemBtn.is_enabled()) == True :
                #退勤ボタン押下
                element.click()
                
                #出勤ボタンが表示されるまで待機
                wait = WebDriverWait(driver, 600)
                selectorS = '//*[@id="container"]/section/main/div[2]/div[2]/ul/li[1]/button'
                element = wait.until(EC.visibility_of_element_located((By.XPATH, selectorS)))
                #出勤ボタンの要素を取得
                elemBtn = driver.find_element(By.XPATH, selectorS)
                #出勤ボタンが押せる場合、ブラウザを閉じ3病後にwindowsシャットダウン。
                if (elemBtn.is_enabled() == True):
                    #ブラウザを終了する。
                    driver.quit()
                    #windows 3秒後に終了
                    os.system(f'shutdown /s /t 3')    
            else: #すでに退勤ボタンが押された状態でしたのポップアップを表示し、jinjar開いたままにする。
                pass
                #退勤ボタンが押せ無い状態（すでに「退勤」ボタンが押された状態。）
                #jinjar確認するか（する「はい」→jinjar開いたまま、しない「いいえ」→jinjar閉じる）
                ansCloseUpDriver = closeUpDriver('退勤')
                if ansCloseUpDriver == 'yes':
                    #ブラウザを終了する。
                    driver.quit()
                else:
                    pass
        elif ansAskQuestion == False:
            #いいえ選択時
            #print('いいえ選択。退勤打刻しません。シャットダウン処理を開始します。')
            # # Edgeを使う場合
            # driver = webdriver.Edge(EdgeChromiumDriverManager().install())
            # #ブラウザを終了する。
            # driver.quit()
            #windows 3秒後に終了
            os.system(f'shutdown /s /t 3')
        else:
            #キャンセル選択時
            #print('キャンセル選択。退勤打刻しません。シャットダウン処理しません。')   
            # # Edgeを使う場合
            # driver = webdriver.Edge(EdgeChromiumDriverManager().install())
            # #ブラウザを終了する。
            # driver.quit()
            pass
except:
    # メッセージボックス（エラー） 
    messagebox.showerror('エラー', 'エラーが発生しました。自動打刻処理を終了します。')
    #ブラウザを終了する。
    driver.quit()