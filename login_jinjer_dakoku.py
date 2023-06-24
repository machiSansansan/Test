from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from time import sleep
from tkinter import messagebox
import os
import signal
import sys
import tkinter as tk
import yaml
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options

root = None
resAskQuestion = None
resAnsRunningAppSearch = None

#ブラウザEdgeを指定
browser = 'executable_path="msedgedriver.exe"'
driver = webdriver.Edge(browser)

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

def runningAppSearch():
    # 起動中のアプリ確認メッセージボックス表示
    resAnsRunningAppSearch = messagebox.askquestion("起動中のアプリ確認", "起動中のアプリはありませんか? \n \n \
    「はい」: 起動中のアプリが無い場合選択。 \n \
    「いいえ」 : 起動中のアプリがあった場合選択。")
    generateWindow()
    return resAnsRunningAppSearch

def getAskQuestion():
    #退勤打刻確認メッセージボックス表示
    resAskQuestion = messagebox.askquestion("jinjer退勤打刻確認", "退勤打刻しますか? \n \n \
    「はい」: 打刻後、windowsシャットダウン。 \n \
    「いいえ」 : windowsシャットダウン。")
    generateWindow()
    return resAskQuestion

def loginInfo():
    url = "https://kintai.jinjer.biz/sign_in"
    driver.get(url)
    
    loginCode = driver.find_element(By.NAME,'company_code')
    loginEmail = driver.find_element(By.NAME,'email')
    loginPass = driver.find_element(By.NAME,'password')
    
    loginCode.send_keys(CODE)
    loginEmail.send_keys(MAIL)
    loginPass.send_keys(PASS)

    btn = driver.find_element(By.NAME,'button')

    #2秒待機
    sleep(2)

    btn.click()
    
    #2秒待機
    sleep(2)

try:
    #出勤退勤フラグ(引数が指定されている場合のみ)
    taikinFlg = '0'
    if len(sys.argv) > 1:
        taikinFlg = sys.argv[1]

    #スタートアップからの起動時
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
            #2秒待機
            sleep(2)
            #出勤ボタン押下
            element.click()
            #2秒待機
            sleep(2)
            # ログアウトする。「ログイン者名」クリック→プルダウン表示された「ログアウト」クリック
            driver.find_element(By.XPATH,'/html/body/header/div/div[3]/ul/li[5]/a').click()
            #2秒待機
            sleep(2)
            driver.find_element(By.XPATH,'/html/body/header/div/div[3]/ul/li[5]/ul/li/a').click()
            #2秒待機
            sleep(2)
            #ブラウザを終了する。
            driver.close()
    else:
        #起動中アプリ確認
        ansRunningAppSearch = runningAppSearch()
        if ansRunningAppSearch == 'yes':
            print('はい選択。起動中アプリはありません。打刻処理するか確認')
            #打刻処理実行
            ansAskQuestion = getAskQuestion()
            if ansAskQuestion == 'yes':
                print('はい選択。退勤打刻します。打刻後、シャットダウン')
                #ログイン処理　開始
                loginInfo()
                #退勤ボタンが表示されるまで待機
                wait = WebDriverWait(driver, 600)
                selectorE = '/html/body/div[1]/div[4]/div[1]/section/main/div[2]/div[2]/ul/li[2]/button'
                element = wait.until(EC.visibility_of_element_located((By.XPATH, selectorE)))

                #退勤ボタンの要素を取得
                elemBtn = driver.find_element(By.XPATH, selectorE)

                # 退勤ボタンが押せる場合、ポップアップを表示し、jinjar開いたままにする。
                if (elemBtn.is_enabled()) == True :
                    #2秒待機
                    sleep(2)
                    #退勤ボタン押下
                    element.click()
                    #2秒待機
                    sleep(2)
                    # ログアウトする。「ログイン者名」クリック→プルダウン表示された「ログアウト」クリック
                    driver.find_element(By.XPATH,'/html/body/header/div/div[3]/ul/li[5]/a').click()
                    #2秒待機
                    sleep(2)
                    driver.find_element(By.XPATH,'/html/body/header/div/div[3]/ul/li[5]/ul/li/a').click()
                    #2秒待機
                    sleep(2)
                    #ブラウザを終了する。
                    driver.close()
                    #windows 5秒後に終了
                    os.system(f'shutdown /s /t 5')
            else:
                print('いいえ選択。退勤打刻しません。シャットダウン処理を開始します。')
                #ブラウザを終了する。
                driver.close()
                #windows 5秒後に終了
                os.system(f'shutdown /s /t 5')
        else:
            print('いいえ選択。起動中アプリがありました。シャットダウン処理を中断します。')
            #ブラウザを終了する。
            driver.close()
finally:
    os.kill(driver.service.process.pid,signal.SIGTERM)