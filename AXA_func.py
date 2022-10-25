#!/usr/bin/env python
# coding: utf-8

# In[1]:



from selenium import webdriver
from time import sleep
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.keys import Keys

import pandas as pd
import openpyxl as xl
import re


# In[2]:


def AXA_func(data):

        #dict型のdataを受け取り、打鍵結果を入力したdict型のdataを返す
    try:
        options = webdriver.ChromeOptions()
        #options.add_argument('--headless') #ブラウザ表示なし
        options.add_argument('--incognito') #シークレットモード 
        options.add_experimental_option('debuggerAddress','127.0.0.1:9222') #アクサはSeleniumから起動すると画面が表示されない。
        browser = webdriver.Chrome(options=options)

        browser.switch_to.new_window('tab')

        url= "https://www.axa-direct.co.jp/qb/html/#/baseTransactionWizard/ContractType"
        browser.get(url)


        if 'S' in str(data['NF等級']):
            browser.find_element(By.CSS_SELECTOR, '#page-inner > div > div > main > div > div.gw-content-wrapper > div > div > div > div.gw-wizard-main > div.gw-page.gw-box.ng-isolate-scope > div:nth-child(2) > div > ng-form > div > div > main > div > div > div.adjEntrance-body-wrapper > div > div:nth-child(3) > div.adjEntrance-item.ng-isolate-scope > div > div:nth-child(2) > ng-transclude > ul > li:nth-child(1) > label > span').click()#他社移行
            #始期日
            Select(browser.find_element(By.CSS_SELECTOR, '#page-inner > div > div > main > div > div.gw-content-wrapper > div > div > div > div.gw-wizard-main > div.gw-page.gw-box.ng-isolate-scope > div:nth-child(2) > div > ng-form > div > div > main > div > div > div.adjEntrance-body-wrapper > div > div:nth-child(3) > div.adjEntrance-block.adjEntrance-start-date.ng-isolate-scope > div > div:nth-child(2) > ng-transclude > div > div.adjInputted-item-select-container.adjInputted-item-select-container-year > select')).select_by_visible_text(str(data['年2'])) 
            Select(browser.find_element(By.CSS_SELECTOR, '#page-inner > div > div > main > div > div.gw-content-wrapper > div > div > div > div.gw-wizard-main > div.gw-page.gw-box.ng-isolate-scope > div:nth-child(2) > div > ng-form > div > div > main > div > div > div.adjEntrance-body-wrapper > div > div:nth-child(3) > div.adjEntrance-block.adjEntrance-start-date.ng-isolate-scope > div > div:nth-child(2) > ng-transclude > div > div.adjInputted-item-select-container.adjInputted-item-select-container-month > select')).select_by_visible_text(str(data['月2']))
            Select(browser.find_element(By.CSS_SELECTOR, '#page-inner > div > div > main > div > div.gw-content-wrapper > div > div > div > div.gw-wizard-main > div.gw-page.gw-box.ng-isolate-scope > div:nth-child(2) > div > ng-form > div > div > main > div > div > div.adjEntrance-body-wrapper > div > div:nth-child(3) > div.adjEntrance-block.adjEntrance-start-date.ng-isolate-scope > div > div:nth-child(2) > ng-transclude > div > div.adjInputted-item-select-container.adjInputted-item-select-container-day > select')).select_by_visible_text(str(data['日2']))
        else:
            browser.find_element(By.XPATH, '//*[@id="page-inner"]/div/div/main/div/div[3]/div/div/div/div[3]/div[1]/div[2]/div/ng-form/div/div/main/div/div/div[1]/div/div[3]/div[1]/div/div[2]/ng-transclude/ul/li[2]/label').click()#他社移行

        browser.find_element(By.XPATH, '//*[@id="direction"]/ul/li[2]/button/span[1]').click()#次へ



        ###保険契約について###########################################################


        if 'S' in str(data['NF等級']):
            pass
        else:
            #ページ遷移完了まで待機、一定時間でタイムアウト
            for t in range(120):
                if len(browser.find_elements(By.XPATH, '//*[@id="section01"]/div[2]/ng-transclude/div/div[1]/div/div/div/ng-transclude/div/ul/li[2]/label/span'))>0:
                    sleep(1)#念の為一秒待ってからつぎへ。
                    break
                else:
                    sleep(1)
            
            browser.find_element(By.XPATH, '//*[@id="section01"]/div[2]/ng-transclude/div/div[1]/div/div/div/ng-transclude/div/ul/li[2]/label/span').click()#東京海上

            Select(browser.find_element(By.XPATH, '//*[@id="previousPolicyTerm_Adj"]')).select_by_visible_text('1年') #1年契約

            #満期日
            Select(browser.find_element(By.XPATH, '//*[@id="section02"]/div[2]/ng-transclude/div/div[2]/div[2]/div/div[2]/ng-transclude/div/div[1]/select')).select_by_visible_text(str(data['年2'])) 
            Select(browser.find_element(By.XPATH, '//*[@id="section02"]/div[2]/ng-transclude/div/div[2]/div[2]/div/div[2]/ng-transclude/div/div[2]/select')).select_by_visible_text(str(data['月2']))
            Select(browser.find_element(By.XPATH, '//*[@id="section02"]/div[2]/ng-transclude/div/div[2]/div[2]/div/div[2]/ng-transclude/div/div[3]/select')).select_by_visible_text(str(data['日2']))

            Select(browser.find_element(By.XPATH, '//*[@id="section03"]/div[2]/ng-transclude/div/div[1]/div/div/div/ng-transclude/div[2]/div/select')).select_by_visible_text(data['NF等級2']) #等級

            if data['背番号2']=='なし':
                    Select(browser.find_element(By.XPATH, '//*[@id="section04"]/div[2]/ng-transclude/div/div[1]/div/div/div/ng-transclude/div/div/select')).select_by_visible_text('0年') #事故有係数適用期間
            else:
                    Select(browser.find_element(By.XPATH, '//*[@id="section04"]/div[2]/ng-transclude/div/div[1]/div/div/div/ng-transclude/div/div/select')).select_by_visible_text('2年') #事故有係数適用期間


            browser.find_element(By.XPATH, '//*[@id="section05"]/div[2]/ng-transclude/div/div/div/div/div/ng-transclude/ul/li[1]/label/span').click()#事故なし
            browser.find_element(By.XPATH, '//*[@id="section06"]/div[2]/ng-transclude/div/div/div/div/div/ng-transclude/ul/li[1]/label/span').click()#車両あり

            sleep(3)
            #次へ
            browser.find_elements(By.XPATH, '//*[@id="direction"]/ul/li[2]/button')[0].click()




        #######車について###########################################################
        #ページ遷移完了まで待機、一定時間でタイムアウト
        for t in range(120):
            if len(browser.find_elements(By.XPATH, '//*[@id="select07"]'))>0:
                sleep(1)#念の為一秒待ってからつぎへ。
                break
            else:
                sleep(1)


        #初度登録
        Select(browser.find_element(By.XPATH, '//*[@id="select07"]')).select_by_visible_text(data['初度年2']) 
        Select(browser.find_element(By.XPATH, '//*[@id="firstRegistrationDate"]/div[2]/select')).select_by_visible_text(data['初度月2'])

        #型式
        browser.find_element(By.CSS_SELECTOR, '#input-Model-No').send_keys(data['型式2'])
        sleep(1)
        browser.find_elements(By.XPATH, '//*[@id="model"]/div[1]/div/div/ng-transclude/div[2]/div[2]/button')[0].click()#検索
        sleep(3)
        try:
            Select(browser.find_element(By.XPATH, '//*[@id="VehicleModelCodeSelect"]')).select_by_value('1')#複数選択があれば一番上を選ぶ
        except:
            pass

        #走行距離
        Select(browser.find_element(By.XPATH, '//*[@id="model.code"]')).select_by_value(str(data['走行距離2']))

        #目的
        if data['使用目的2'] == '日常':
            browser.find_element(By.XPATH, '//*[@id="section03"]/div[2]/ng-transclude/div[2]/div/div/div/div[2]/ng-transclude/div/ul/li[1]/label').click()#日常
        elif data['使用目的2'] == '通勤':
            browser.find_element(By.XPATH, '//*[@id="section03"]/div[2]/ng-transclude/div[2]/div/div/div/div[2]/ng-transclude/div/ul/li[2]/label').click()#通勤・通学
        else:
            browser.find_element(By.XPATH, '//*[@id="section03"]/div[2]/ng-transclude/div[2]/div/div/div/div[2]/ng-transclude/div/ul/li[3]/label').click()#業務

        #子育て応援割り
        browser.find_element(By.XPATH, '//*[@id="section03"]/div[2]/ng-transclude/div[5]/div/div[2]/ng-transclude/adj-yes-no/ul/li[2]/label').click()#なし

        #セカンドカー割引
        if str(data['NF等級']) =='7S':
            browser.find_element(By.XPATH, '//*[@id="section04"]/div[2]/ng-transclude/div[1]/div/adj-yes-no/ul/li[1]/label/span').click()#あり
            sleep(1)
            browser.find_element(By.XPATH, '//*[@id="vehicleUseType"]/div[1]/div/div[2]/ng-transclude/div/ul/li[1]/label').click()#自普乗
            Select(browser.find_element(By.XPATH, '//*[@id="insurance"]/div/div/div/div/div[2]/ng-transclude/div/div/select')).select_by_value('1')#他社
            Select(browser.find_element(By.XPATH, '//*[@id="grade"]/div/div/div/div/div[2]/ng-transclude/div/div/select')).select_by_value('1')#等級
            
            #年をどう選ぶか検討　月日は1/1固定
            Select(browser.find_element(By.XPATH, '//*[@id="first-registrant"]/div/div/div/div/div[2]/ng-transclude/div/div[1]/select')).select_by_index(8)
            Select(browser.find_element(By.XPATH, '//*[@id="first-registrant"]/div/div/div/div/div[2]/ng-transclude/div/div[2]/select')).select_by_value('number:0')
            Select(browser.find_element(By.XPATH, '//*[@id="first-registrant"]/div/div/div/div/div[2]/ng-transclude/div/div[3]/select')).select_by_value('number:1')
            #年をどう選ぶか検討　月日は1/1固定
            Select(browser.find_element(By.XPATH, '//*[@id="expiry"]/div/div/div/div/div[2]/ng-transclude/div/div[1]/select')).select_by_index(3)
            Select(browser.find_element(By.XPATH, '//*[@id="expiry"]/div/div/div/div/div[2]/ng-transclude/div/div[2]/select')).select_by_value('number:0')
            Select(browser.find_element(By.XPATH, '//*[@id="expiry"]/div/div/div/div/div[2]/ng-transclude/div/div[3]/select')).select_by_value('number:1')

            browser.find_element(By.XPATH, '//*[@id="user"]/div/div/div/div/div[2]/ng-transclude/div/ul/li[1]/label').click()
            browser.find_element(By.XPATH, '//*[@id="driver"]/div/div/div/div/div[2]/ng-transclude/div/ul/li[1]/label').click()
        elif str(data['NF等級']) =='6S':
            browser.find_element(By.XPATH, '//*[@id="section04"]/div[2]/ng-transclude/div[1]/div/adj-yes-no/ul/li[2]/label/span').click()#なし
        else:
            pass


        sleep(1)
        #次へ
        browser.find_element(By.XPATH, '//*[@id="direction"]/ul/li[2]/button').click()#なし


        #####記名被保険者について#########################################################
        #ページ遷移完了まで待機、一定時間でタイムアウト
        for t in range(120):
            if len(browser.find_elements(By.XPATH, '//*[@id="prefecture_Adj"]'))>0:
                sleep(1)#念の為一秒待ってからつぎへ。
                break
            else:
                sleep(1)

        #都道府県
        Select(browser.find_element(By.XPATH, '//*[@id="prefecture_Adj"]')).select_by_visible_text(data['地域2']) 

        #生年月日
        Select(browser.find_element(By.XPATH, '//*[@id="birthdaySection"]/div[2]/ng-transclude/div/div/div/div/ng-transclude/div/div[1]/select')).select_by_visible_text(str(data['生年2'])) 
        Select(browser.find_element(By.XPATH, '//*[@id="birthdaySection"]/div[2]/ng-transclude/div/div/div/div/ng-transclude/div/div[2]/select')).select_by_visible_text(str(data['生まれ月2']))
        Select(browser.find_element(By.XPATH, '//*[@id="birthdaySection"]/div[2]/ng-transclude/div/div/div/div/ng-transclude/div/div[3]/select')).select_by_visible_text(str(data['生まれ日2']))

        #免許の色
        if data['免許2'] == 'ゴールド':
            browser.find_element(By.XPATH, '//*[@id="licenseColorSection"]/div[2]/ng-transclude/div[1]/div/div/div/ng-transclude/div/div[1]/ul/li[1]/label').click()#ゴールド
        elif data['免許2'] == 'ブルー':
            browser.find_element(By.XPATH, '//*[@id="licenseColorSection"]/div[2]/ng-transclude/div[1]/div/div/div/ng-transclude/div/div[1]/ul/li[2]/label').click()#ブルー
        else:    
            browser.find_element(By.XPATH, '//*[@id="licenseColorSection"]/div[2]/ng-transclude/div[1]/div/div/div/ng-transclude/div/div[1]/ul/li[3]/label').click()#グリーン
                                        
        #運転者限定
        if data['運限'] == '本配':
            browser.find_element(By.XPATH, '//*[@id="rangeOfDriversSection"]/div[2]/ng-transclude/div[1]/div/div/div/div[2]/ng-transclude/ul/li[1]/label').click()#本人配偶者
        else:
            browser.find_element(By.XPATH, '//*[@id="rangeOfDriversSection"]/div[2]/ng-transclude/div[1]/div/div/div/div[2]/ng-transclude/ul/li[2]/label').click()#限定なし


        #年齢条件
        if data['年限修正2'] == '全年齢補償':
            browser.find_element(By.CSS_SELECTOR, '#driver > div > div > div > div:nth-child(2) > ng-transclude > ul > li:nth-child(1) > label').click()
        elif data['年限修正2'] == '21歳以上補償':
            browser.find_element(By.CSS_SELECTOR, '#driver > div > div > div > div:nth-child(2) > ng-transclude > ul > li:nth-child(2) > label').click()
        elif data['年限修正2'] == '26歳以上補償':
            browser.find_element(By.CSS_SELECTOR, '#driver > div > div > div > div:nth-child(2) > ng-transclude > ul > li:nth-child(3) > label').click()
        else: # data['年限修正2'] == '30歳以上':
            browser.find_element(By.CSS_SELECTOR, '#driver > div > div > div > div:nth-child(2) > ng-transclude > ul > li:nth-child(4) > label').click()

        sleep(1)


        #次へ
        browser.find_element(By.XPATH, '//*[@id="direction"]/ul/li[2]/button').click()

        ######見積もり画面#################################################################
        #ページ遷移完了まで待機、一定時間でタイムアウト
        for t in range(120):
            if len(browser.find_elements(By.CLASS_NAME, 'adjTariffPcPlan__summary-customize-button'))>0:
                sleep(1)#念の為一秒待ってからつぎへ。
                break
            else:
                sleep(1)


        ########デフォルトの状態での契約条件（特約有無など）を確認、のちにトグルスイッチのクリック要否判定に使う。

        #対物LL、「なし」表示の有無で判定
        P_LL=list(range(2))
        P_LL[0]=len(browser.find_elements(By.CSS_SELECTOR,'#default_0_vehicle_adj\.quoteandbind\.jppa\.directives\.JPPAQuoteDetailsCtrl\.Coverage\ for\ other\ person\ \&\ property_group > div:nth-child(7) > div > div > div > div.adjTariffPcContent__body > div > div:nth-child(1) > div > gw-qnb-quote-view-cell > div > div:nth-child(1) > div > div.adjTariffPcItem.is-negative.ng-scope > p'))
        P_LL[0]=P_LL[0]*(-1)+1
        P_LL[1]=len(browser.find_elements(By.CSS_SELECTOR,'#default_0_vehicle_adj\.quoteandbind\.jppa\.directives\.JPPAQuoteDetailsCtrl\.Coverage\ for\ other\ person\ \&\ property_group > div:nth-child(7) > div > div > div > div.adjTariffPcContent__body > div > div:nth-child(2) > div > gw-qnb-quote-view-cell > div > div:nth-child(1) > div > div.adjTariffPcItem.is-negative.ng-scope > p'))
        P_LL[1]=P_LL[1]*(-1)+1

        #搭乗者傷害は、金額の有無で判定
        P_tosho=list(range(2))
        P_tosho[0]=len(browser.find_elements(By.XPATH,'//*[@id="default_1_vehicle_adj.quoteandbind.jppa.directives.JPPAQuoteDetailsCtrl.Coverage for passengers_group"]/div[9]/div/div/div/div[2]/div/div[1]/div/gw-qnb-quote-view-cell/div/div[2]/div[1]/div[2]/gw-pc-coverage-term/div/p'))
        P_tosho[1]=len(browser.find_elements(By.XPATH,'//*[@id="default_1_vehicle_adj.quoteandbind.jppa.directives.JPPAQuoteDetailsCtrl.Coverage for passengers_group"]/div[9]/div/div/div/div[2]/div/div[2]/div/gw-qnb-quote-view-cell/div/div[2]/div[1]/div[2]/gw-pc-coverage-term/div/p'))

        #地噴津の有無
        P_jifuntsu=list(range(2))
        P_jifuntsu[0]=len(browser.find_elements(By.XPATH,'//*[@id="default_2_vehicle_adj.quoteandbind.jppa.directives.JPPAQuoteDetailsCtrl.Coverage for vehicle damage_group"]/div[5]/div/div/div/div[2]/div/div[1]/div/gw-qnb-quote-view-cell/div/div[2]/div/div[2]/gw-pc-coverage-term/div/p'))
        P_jifuntsu[1]=len(browser.find_elements(By.XPATH,'//*[@id="default_2_vehicle_adj.quoteandbind.jppa.directives.JPPAQuoteDetailsCtrl.Coverage for vehicle damage_group"]/div[5]/div/div/div/div[2]/div/div[2]/div/gw-qnb-quote-view-cell/div/div[2]/div/div[2]/gw-pc-coverage-term/div/p'))

        #弁護士特約
        P_bengoshi=list(range(2))
        P_bengoshi[0]=int((browser.find_element(By.XPATH,'//*[@id="default_3_vehicle_adj.quoteandbind.jppa.directives.JPPAQuoteDetailsCtrl.Other coverages_group"]/div[3]/div/div/div/div[2]/div/div[1]/div/gw-qnb-quote-view-cell/div/div/div/div[2]/p').text) == 'あり')
        P_bengoshi[1]=int((browser.find_element(By.XPATH,'//*[@id="default_3_vehicle_adj.quoteandbind.jppa.directives.JPPAQuoteDetailsCtrl.Other coverages_group"]/div[3]/div/div/div/div[2]/div/div[2]/div/gw-qnb-quote-view-cell/div/div/div/div[2]/p').text) == 'あり')

        #賠償
        P_kobai=list(range(2))
        P_kobai[0]=int((browser.find_element(By.XPATH,'//*[@id="default_3_vehicle_adj.quoteandbind.jppa.directives.JPPAQuoteDetailsCtrl.Other coverages_group"]/div[5]/div/div/div/div[2]/div/div[1]/div/gw-qnb-quote-view-cell/div/div/div/div[2]/p').text) != 'なし')
        P_kobai[1]=int((browser.find_element(By.XPATH,'//*[@id="default_3_vehicle_adj.quoteandbind.jppa.directives.JPPAQuoteDetailsCtrl.Other coverages_group"]/div[5]/div/div/div/div[2]/div/div[2]/div/gw-qnb-quote-view-cell/div/div/div/div[2]/p').text) != 'なし')

        #FB特約
        P_FB=list(range(2))
        P_FB[0]=int((browser.find_element(By.XPATH,'//*[@id="default_3_vehicle_adj.quoteandbind.jppa.directives.JPPAQuoteDetailsCtrl.Other coverages_group"]/div[13]/div[1]/div/div/div[2]/div/div[1]/div/gw-qnb-quote-view-cell/div/div[1]/div/div/p').text) == 'あり')
        P_FB[1]=int((browser.find_element(By.XPATH,'//*[@id="default_3_vehicle_adj.quoteandbind.jppa.directives.JPPAQuoteDetailsCtrl.Other coverages_group"]/div[13]/div[1]/div/div/div[2]/div/div[2]/div/gw-qnb-quote-view-cell/div/div[1]/div/div[2]/p').text) == 'あり')



        #####################プラン１変更
        browser.find_elements(By.CLASS_NAME, 'adjTariffPcPlan__summary-customize-button')[0].click()
        browser.execute_script("window.scrollTo(0, 500);")
        sleep(1)

        #対物全損時修理差額 
        if data['対物LL'] == 'なし':
            if P_LL[0] == 1:
                browser.find_element(By.CSS_SELECTOR, '#default_0_vehicle_adj\.quoteandbind\.jppa\.directives\.JPPAQuoteDetailsCtrl\.Coverage\ for\ other\ person\ \&\ property_group > div:nth-child(7) > div > div > div > div.adjTariffPcContent__body > div > div:nth-child(1) > div > gw-qnb-quote-view-cell > div > div:nth-child(1) > div > div > div > div > label > span').click()
        else:
            if P_LL[0] == 0:
                browser.find_element(By.CSS_SELECTOR, '#default_0_vehicle_adj\.quoteandbind\.jppa\.directives\.JPPAQuoteDetailsCtrl\.Coverage\ for\ other\ person\ \&\ property_group > div:nth-child(7) > div > div > div > div.adjTariffPcContent__body > div > div:nth-child(1) > div > gw-qnb-quote-view-cell > div > div:nth-child(1) > div > div > div > div > label > span').click()
        sleep(1)


        #人傷内外
        Select(browser.find_element(By.CSS_SELECTOR, '#JPPAPersonalInjuryLimitCovTerm_Adj')).select_by_visible_text(data['人傷AMT2'])
        sleep(1)
        browser.find_element(By.CSS_SELECTOR, '#default_1_vehicle_adj\.quoteandbind\.jppa\.directives\.JPPAQuoteDetailsCtrl\.Coverage\ for\ passengers_group > div:nth-child(3) > div > div > div > div.adjTariffPcContent__body > div > div:nth-child(1) > div > gw-qnb-quote-view-cell > div > div:nth-child(2) > div:nth-child(2) > div > gw-pc-coverage-term > div > div > div > div.adjEstItem__body > div > div > div:nth-child(2) > label > div').click()#搭乗中のみ、なし
        sleep(1)

        #搭乗者
        if data['搭傷2'] != 'なし':
            if P_tosho[0] == 0:
                browser.find_element(By.CSS_SELECTOR, '#default_1_vehicle_adj\.quoteandbind\.jppa\.directives\.JPPAQuoteDetailsCtrl\.Coverage\ for\ passengers_group > div:nth-child(9) > div > div > div > div.adjTariffPcContent__body > div > div:nth-child(1) > div > gw-qnb-quote-view-cell > div > div:nth-child(1) > div > div > div > div > label > span').click()#搭傷ありに変更
                sleep(1)
                Select(browser.find_element(By.CSS_SELECTOR, '#JPPADPPADeathDisabilityLimitCovTerm_Adj')).select_by_visible_text(data['搭傷2'])
            else:
                sleep(1)
                Select(browser.find_element(By.CSS_SELECTOR, '#JPPADPPADeathDisabilityLimitCovTerm_Adj')).select_by_visible_text(data['搭傷2'])
        else:
            if P_tosho[0] == 1:
                browser.find_element(By.XPATH, '//*[@id="default_1_vehicle_adj.quoteandbind.jppa.directives.JPPAQuoteDetailsCtrl.Coverage for passengers_group"]/div[9]/div/div/div/div[2]/div/div[1]/div/gw-qnb-quote-view-cell/div/div[1]/div/div/div/div/label/span').click()#搭傷なしに変更
        sleep(2)



        #車両保険
        #if len(browser.find_elements(By.CSS_SELECTOR, '#JPPAOwnDamageCompensationTypeCovTerm_Adj')) ==0:
        #    sleep(1)
        #browser.find_element(By.CSS_SELECTOR, '#default_2_vehicle_adj\.quoteandbind\.jppa\.directives\.JPPAQuoteDetailsCtrl\.Coverage\ for\ vehicle\ damage_group > div:nth-child(3) > div > div > div > div.adjTariffPcContent__body > div > div:nth-child(1) > div > gw-qnb-quote-view-cell > div > div:nth-child(1) > div > div > div > div > label > span').click()
        sleep(1)
        Select(browser.find_element(By.CSS_SELECTOR, '#JPPAOwnDamageCompensationTypeCovTerm_Adj')).select_by_visible_text('一般車両保険')
        sleep(1)
        #改行が入っていてテキストで選べないためindexをつかう
        data['免責2']=data['免責2'].replace('1回目事故 5万円（車対車免ゼロ特約付）　 2回目以降事故 10万円','0')
        data['免責2']=data['免責2'].replace('1回目事故 5万円　 2回目以降事故 10万円','1')
        data['免責2']=data['免責2'].replace('1回目事故 0万円　 2回目以降事故 10万円','2')
        data['免責2']=data['免責2'].replace('1回目事故 10万円　 2回目以降事故 10万円','3')
        Select(browser.find_element(By.CSS_SELECTOR, '#JPPAOwnDamageDeductibleCovTerm_Adj')).select_by_index(int(data['免責2']))
        sleep(1)
        Select(browser.find_element(By.CSS_SELECTOR, '#default_2_vehicle_adj\.quoteandbind\.jppa\.directives\.JPPAQuoteDetailsCtrl\.Coverage\ for\ vehicle\ damage_group > div:nth-child(3) > div > div > div > div.adjTariffPcContent__body > div > div:nth-child(1) > div > gw-qnb-quote-view-cell > div > div:nth-child(2) > div:nth-child(3) > div > gw-pc-coverage-term > div > div > div > div.adjEstItem__body > div > select')).select_by_visible_text(data['車両AMT2'])
        sleep(1)


        #弁特
        if data['弁特2'] == 'あり':
            if P_bengoshi[0] == 0:
                browser.find_element(By.CSS_SELECTOR, '#default_3_vehicle_adj\.quoteandbind\.jppa\.directives\.JPPAQuoteDetailsCtrl\.Other\ coverages_group > div:nth-child(3) > div > div > div > div.adjTariffPcContent__body > div > div:nth-child(1) > div > gw-qnb-quote-view-cell > div > div > div > div > div > div > label > span').click()#ありに変更
        else:
            if P_bengoshi[0] == 1:
                browser.find_element(By.CSS_SELECTOR, '#default_3_vehicle_adj\.quoteandbind\.jppa\.directives\.JPPAQuoteDetailsCtrl\.Other\ coverages_group > div:nth-child(3) > div > div > div > div.adjTariffPcContent__body > div > div:nth-child(1) > div > gw-qnb-quote-view-cell > div > div > div > div > div > div > label > span').click()#なしに変更
        sleep(1)

        #個賠
        if data['個賠'] == 'あり':
            if P_kobai[0] == 0:
                browser.find_element(By.CSS_SELECTOR, '#default_3_vehicle_adj\.quoteandbind\.jppa\.directives\.JPPAQuoteDetailsCtrl\.Other\ coverages_group > div:nth-child(5) > div > div > div > div.adjTariffPcContent__body > div > div:nth-child(1) > div > gw-qnb-quote-view-cell > div > div > div > div > div > div > label > span').click()#ありに変更
        else:
            if P_kobai[0] == 1:
                browser.find_element(By.CSS_SELECTOR, '#default_3_vehicle_adj\.quoteandbind\.jppa\.directives\.JPPAQuoteDetailsCtrl\.Other\ coverages_group > div:nth-child(5) > div > div > div > div.adjTariffPcContent__body > div > div:nth-child(1) > div > gw-qnb-quote-view-cell > div > div > div > div > div > div > label > span').click()#なしに変更
        sleep(1)

        #ファミリープラス
        if data['ファミリー'] == 'あり':
            browser.find_element(By.CSS_SELECTOR, '#default_3_vehicle_adj\.quoteandbind\.jppa\.directives\.JPPAQuoteDetailsCtrl\.Other\ coverages_group > div:nth-child(7) > div > div > div.adjTariffPcContent__body > div > div:nth-child(1) > div > div > div:nth-child(1) > div > label > span').click()#ありに変更
            if len(browser.find_elements(By.CSS_SELECTOR, 'body > div.gw-modal.gw-fade.gw-modal_animation_.in > div.adjContainer.gw-modal__inner > div > div > div > div > button'))>0:
                browser.find_element(By.CSS_SELECTOR, 'body > div.gw-modal.gw-fade.gw-modal_animation_.in > div.adjContainer.gw-modal__inner > div > div > div > div > button').click()#閉じるボタン
        sleep(1)

        #レディース
        if data['レディース'] == 'あり':
            browser.find_element(By.CSS_SELECTOR, '#default_3_vehicle_adj\.quoteandbind\.jppa\.directives\.JPPAQuoteDetailsCtrl\.Other\ coverages_group > div:nth-child(7) > div > div > div.adjTariffPcContent__body > div > div:nth-child(1) > div > div > div:nth-child(2) > div > label > span').click()#ありに変更
            if len(browser.find_elements(By.CSS_SELECTOR, 'body > div.gw-modal.gw-fade.gw-modal_animation_.in > div.adjContainer.gw-modal__inner > div > div > div > div > button'))>0:
                browser.find_element(By.CSS_SELECTOR, 'body > div.gw-modal.gw-fade.gw-modal_animation_.in > div.adjContainer.gw-modal__inner > div > div > div > div > button').click()#閉じるボタン
        sleep(1)

        #ペット
        if data['ペット'] == 'あり':
            browser.find_element(By.CSS_SELECTOR, '#default_3_vehicle_adj\.quoteandbind\.jppa\.directives\.JPPAQuoteDetailsCtrl\.Other\ coverages_group > div:nth-child(7) > div > div > div.adjTariffPcContent__body > div > div:nth-child(1) > div > div > div:nth-child(3) > div > label > span').click()#ありに変更
            if len(browser.find_elements(By.CSS_SELECTOR, 'body > div.gw-modal.gw-fade.gw-modal_animation_.in > div.adjContainer.gw-modal__inner > div > div > div > div > button'))>0:
                browser.find_element(By.CSS_SELECTOR, 'body > div.gw-modal.gw-fade.gw-modal_animation_.in > div.adjContainer.gw-modal__inner > div > div > div > div > button').click()#閉じるボタン
        sleep(1)

        #ファミリーバイク
        if data['ファミリー2'] == 'あり':
            if P_FB[0] == 0:
                browser.find_element(By.CSS_SELECTOR, '#default_3_vehicle_adj\.quoteandbind\.jppa\.directives\.JPPAQuoteDetailsCtrl\.Other\ coverages_group > div:nth-child(13) > div:nth-child(1) > div > div > div.adjTariffPcContent__body > div > div:nth-child(1) > div > gw-qnb-quote-view-cell > div > div:nth-child(1) > div > div > div > div > label > span').click()#ありに変更
        else:
            if P_FB[0] == 1:
                browser.find_element(By.CSS_SELECTOR, '#default_3_vehicle_adj\.quoteandbind\.jppa\.directives\.JPPAQuoteDetailsCtrl\.Other\ coverages_group > div:nth-child(13) > div:nth-child(1) > div > div > div.adjTariffPcContent__body > div > div:nth-child(1) > div > gw-qnb-quote-view-cell > div > div:nth-child(1) > div > div > div > div > label > span').click()#なしに変更
        sleep(2)

        #再計算ボタン
        browser.find_elements(By.CLASS_NAME, 'adjTariffPcPlan__summary-apply-button')[0].click()
        sleep(2)


        ######################プラン２変更
        browser.execute_script("window.scrollTo(0, 0)")
        browser.find_elements(By.CLASS_NAME, 'adjTariffPcPlan__summary-customize-button')[2].click()
        browser.execute_script("window.scrollTo(0, 500)")
        sleep(1)

        #対物全損時修理差額 
        if data['対物LL'] == 'なし':
            if P_LL[1] == 1:
                browser.find_element(By.CSS_SELECTOR, '#default_0_vehicle_adj\.quoteandbind\.jppa\.directives\.JPPAQuoteDetailsCtrl\.Coverage\ for\ other\ person\ \&\ property_group > div:nth-child(7) > div > div > div > div.adjTariffPcContent__body > div > div:nth-child(2) > div > gw-qnb-quote-view-cell > div > div:nth-child(1) > div > div > div > div > label').click()
        else:
            if P_LL[1] == 0:
                browser.find_element(By.CSS_SELECTOR, '#default_0_vehicle_adj\.quoteandbind\.jppa\.directives\.JPPAQuoteDetailsCtrl\.Coverage\ for\ other\ person\ \&\ property_group > div:nth-child(7) > div > div > div > div.adjTariffPcContent__body > div > div:nth-child(2) > div > gw-qnb-quote-view-cell > div > div:nth-child(1) > div > div > div > div > label').click()
        sleep(1)

        #人傷内外3000万円
        Select(browser.find_element(By.CSS_SELECTOR, '#JPPAPersonalInjuryLimitCovTerm_Adj')).select_by_visible_text(data['人傷AMT2'])
        sleep(1)
        browser.find_element(By.CSS_SELECTOR, '#default_1_vehicle_adj\.quoteandbind\.jppa\.directives\.JPPAQuoteDetailsCtrl\.Coverage\ for\ passengers_group > div:nth-child(3) > div > div > div > div.adjTariffPcContent__body > div > div:nth-child(2) > div > gw-qnb-quote-view-cell > div > div:nth-child(2) > div:nth-child(2) > div > gw-pc-coverage-term > div > div > div > div.adjEstItem__body > div > div > div:nth-child(2) > label > div').click()#搭乗中のみ、なし
        sleep(1)

        #搭乗者
        if data['搭傷2'] != 'なし':
            if P_tosho[1] == 0:
                browser.find_element(By.CSS_SELECTOR, '#default_1_vehicle_adj\.quoteandbind\.jppa\.directives\.JPPAQuoteDetailsCtrl\.Coverage\ for\ passengers_group > div:nth-child(9) > div > div > div > div.adjTariffPcContent__body > div > div:nth-child(2) > div > gw-qnb-quote-view-cell > div > div:nth-child(1) > div > div > div > div > label > span').click()#搭傷ありに変更
                sleep(1)
                Select(browser.find_element(By.CSS_SELECTOR, '#JPPADPPADeathDisabilityLimitCovTerm_Adj')).select_by_visible_text(data['搭傷2'])
            else:
                sleep(1)
                Select(browser.find_element(By.CSS_SELECTOR, '#JPPADPPADeathDisabilityLimitCovTerm_Adj')).select_by_visible_text(data['搭傷2'])
        else:
            if P_tosho[1] == 1:
                browser.find_element(By.CSS_SELECTOR, '#default_1_vehicle_adj\.quoteandbind\.jppa\.directives\.JPPAQuoteDetailsCtrl\.Coverage\ for\ passengers_group > div:nth-child(9) > div > div > div > div.adjTariffPcContent__body > div > div:nth-child(2) > div > gw-qnb-quote-view-cell > div > div:nth-child(1) > div > div > div > div > label > span').click()#搭傷なしに変更
        sleep(2)

        #車両保険
        browser.find_element(By.XPATH, '//*[@id="default_2_vehicle_adj.quoteandbind.jppa.directives.JPPAQuoteDetailsCtrl.Coverage for vehicle damage_group"]/div[3]/div/div/div/div[3]/div/div[2]/div/gw-qnb-quote-view-cell/div/div[1]/div/div/div/div/label/span').click()#車両なし
        sleep(2)

        #弁特
        if data['弁特2'] == 'あり':
            if P_bengoshi[1] == 0:
                browser.find_element(By.CSS_SELECTOR, '#default_3_vehicle_adj\.quoteandbind\.jppa\.directives\.JPPAQuoteDetailsCtrl\.Other\ coverages_group > div:nth-child(3) > div > div > div > div.adjTariffPcContent__body > div > div:nth-child(2) > div > gw-qnb-quote-view-cell > div > div > div > div > div > div > label > span').click()#ありに変更
        else:
            if P_bengoshi[1] == 1:
                browser.find_element(By.CSS_SELECTOR, '#default_3_vehicle_adj\.quoteandbind\.jppa\.directives\.JPPAQuoteDetailsCtrl\.Other\ coverages_group > div:nth-child(3) > div > div > div > div.adjTariffPcContent__body > div > div:nth-child(2) > div > gw-qnb-quote-view-cell > div > div > div > div > div > div > label > span').click()#なしに変更
        sleep(1)

        #個賠
        if data['個賠'] == 'あり':
            if P_kobai[1] == 0:
                browser.find_element(By.CSS_SELECTOR, '#default_3_vehicle_adj\.quoteandbind\.jppa\.directives\.JPPAQuoteDetailsCtrl\.Other\ coverages_group > div:nth-child(5) > div > div > div > div.adjTariffPcContent__body > div > div:nth-child(2) > div > gw-qnb-quote-view-cell > div > div > div > div > div > div > label > span').click()#ありに変更
        else:
            if P_kobai[1] == 1:
                browser.find_element(By.CSS_SELECTOR, '#default_3_vehicle_adj\.quoteandbind\.jppa\.directives\.JPPAQuoteDetailsCtrl\.Other\ coverages_group > div:nth-child(5) > div > div > div > div.adjTariffPcContent__body > div > div:nth-child(2) > div > gw-qnb-quote-view-cell > div > div > div > div > div > div > label > span').click()#なしに変更
        sleep(1)

        #ファミリープラス
        if data['ファミリー'] == 'あり':
            browser.find_element(By.CSS_SELECTOR, '#default_3_vehicle_adj\.quoteandbind\.jppa\.directives\.JPPAQuoteDetailsCtrl\.Other\ coverages_group > div:nth-child(7) > div > div > div.adjTariffPcContent__body > div > div:nth-child(2) > div > div > div:nth-child(1) > div > label > span').click()#ありに変更
            if len(browser.find_elements(By.CSS_SELECTOR, 'body > div.gw-modal.gw-fade.gw-modal_animation_.in > div.adjContainer.gw-modal__inner > div > div > div > div > button'))>0:
                browser.find_element(By.CSS_SELECTOR, 'body > div.gw-modal.gw-fade.gw-modal_animation_.in > div.adjContainer.gw-modal__inner > div > div > div > div > button').click()#閉じるボタン
        sleep(1)

        #レディース
        if data['レディース'] == 'あり':
            browser.find_element(By.CSS_SELECTOR, '#default_3_vehicle_adj\.quoteandbind\.jppa\.directives\.JPPAQuoteDetailsCtrl\.Other\ coverages_group > div:nth-child(7) > div > div > div.adjTariffPcContent__body > div > div:nth-child(2) > div > div > div:nth-child(2) > div > label > span').click()#ありに変更
            if len(browser.find_elements(By.CSS_SELECTOR, 'body > div.gw-modal.gw-fade.gw-modal_animation_.in > div.adjContainer.gw-modal__inner > div > div > div > div > button'))>0:
                browser.find_element(By.CSS_SELECTOR, 'body > div.gw-modal.gw-fade.gw-modal_animation_.in > div.adjContainer.gw-modal__inner > div > div > div > div > button').click()#閉じるボタン
        sleep(1)

        #ペット
        if data['ペット'] == 'あり':
            browser.find_element(By.CSS_SELECTOR, '#default_3_vehicle_adj\.quoteandbind\.jppa\.directives\.JPPAQuoteDetailsCtrl\.Other\ coverages_group > div:nth-child(7) > div > div > div.adjTariffPcContent__body > div > div:nth-child(2) > div > div > div:nth-child(3) > div > label > span').click()#ありに変更
            if len(browser.find_elements(By.CSS_SELECTOR, 'body > div.gw-modal.gw-fade.gw-modal_animation_.in > div.adjContainer.gw-modal__inner > div > div > div > div > button'))>0:
                browser.find_element(By.CSS_SELECTOR, 'body > div.gw-modal.gw-fade.gw-modal_animation_.in > div.adjContainer.gw-modal__inner > div > div > div > div > button').click()#閉じるボタン
        sleep(1)

        #ファミリーバイク
        if data['ファミリー2'] == 'あり':
            if P_FB[1] == 0:
                browser.find_element(By.CSS_SELECTOR, '#default_3_vehicle_adj\.quoteandbind\.jppa\.directives\.JPPAQuoteDetailsCtrl\.Other\ coverages_group > div:nth-child(13) > div:nth-child(1) > div > div > div.adjTariffPcContent__body > div > div:nth-child(2) > div > gw-qnb-quote-view-cell > div > div:nth-child(1) > div > div > div > div > label > span').click()#ありに変更
        else:
            if P_FB[1] == 1:
                browser.find_element(By.CSS_SELECTOR, '#default_3_vehicle_adj\.quoteandbind\.jppa\.directives\.JPPAQuoteDetailsCtrl\.Other\ coverages_group > div:nth-child(13) > div:nth-child(1) > div > div > div.adjTariffPcContent__body > div > div:nth-child(2) > div > gw-qnb-quote-view-cell > div > div:nth-child(1) > div > div > div > div > label > span').click()#なしに変更
        sleep(2)

        #再計算ボタン
        browser.find_elements(By.CLASS_NAME, 'adjTariffPcPlan__summary-apply-button')[1].click()
        sleep(1)

        browser.execute_script("window.scrollTo(0, 0)")#上までスクロールしないとエラーが起きがち

        ######保険料取得####################################
        result0_discount=browser.find_element(By.CSS_SELECTOR, '#page-inner > div > div > main > div > div.gw-content-wrapper > div > div > div > div.gw-wizard-main > div.gw-page.gw-box.ng-isolate-scope > div:nth-child(2) > div > gw-qnb-custom-quote > form > div > gw-qnb-common-offering-selection > div > gw-qnb-multiple-offering-view > div > div.adjTariffPc__main-container > div.adjTariffPc__main-body > div > div > div > div > div:nth-child(1) > div > div > div.adjTariffPcPlan__summary-price > p.adjTariffPcPlan__summary-price-current > span').text
        result0_discount=int(re.sub(r"\D", "", result0_discount))

        result0=browser.find_element(By.CSS_SELECTOR, '#page-inner > div > div > main > div > div.gw-content-wrapper > div > div > div > div.gw-wizard-main > div.gw-page.gw-box.ng-isolate-scope > div:nth-child(2) > div > gw-qnb-custom-quote > form > div > gw-qnb-common-offering-selection > div > gw-qnb-multiple-offering-view > div > div.adjTariffPc__main-container > div.adjTariffPc__main-body > div > div > div > div > div:nth-child(1) > div > div > div.adjTariffPcPlan__summary-price > p.adjTariffPcPlan__summary-price-before > span').text
        result0=int(re.sub(r"\D", "", result0))


        #ここの表示に時間がかかるため待機、一定時間でタイムアウト
        for t in range(120):
            if browser.find_element(By.CSS_SELECTOR, '#page-inner > div > div > main > div > div.gw-content-wrapper > div > div > div > div.gw-wizard-main > div.gw-page.gw-box.ng-isolate-scope > div:nth-child(2) > div > gw-qnb-custom-quote > form > div > gw-qnb-common-offering-selection > div > gw-qnb-multiple-offering-view > div > div.adjTariffPc__main-container > div.adjTariffPc__main-body > div > div > div > div > div:nth-child(2) > div > div > div.adjTariffPcPlan__summary-price > p.adjTariffPcPlan__summary-price-current > span').text != '- 円' :
                sleep(1)#念の為一秒待ってからつぎへ。
                break
            else:
                sleep(1)

        result1_discount=browser.find_element(By.CSS_SELECTOR, '#page-inner > div > div > main > div > div.gw-content-wrapper > div > div > div > div.gw-wizard-main > div.gw-page.gw-box.ng-isolate-scope > div:nth-child(2) > div > gw-qnb-custom-quote > form > div > gw-qnb-common-offering-selection > div > gw-qnb-multiple-offering-view > div > div.adjTariffPc__main-container > div.adjTariffPc__main-body > div > div > div > div > div:nth-child(2) > div > div > div.adjTariffPcPlan__summary-price > p.adjTariffPcPlan__summary-price-current > span').text
        result1_discount=int(re.sub(r"\D", "", result1_discount))


        result1=browser.find_element(By.CSS_SELECTOR, '#page-inner > div > div > main > div > div.gw-content-wrapper > div > div > div > div.gw-wizard-main > div.gw-page.gw-box.ng-isolate-scope > div:nth-child(2) > div > gw-qnb-custom-quote > form > div > gw-qnb-common-offering-selection > div > gw-qnb-multiple-offering-view > div > div.adjTariffPc__main-container > div.adjTariffPc__main-body > div > div > div > div > div:nth-child(2) > div > div > div.adjTariffPcPlan__summary-price > p.adjTariffPcPlan__summary-price-before > span').text
        result1=int(re.sub(r"\D", "", result1))

        ###賠償
        compensation0=browser.find_element(By.CSS_SELECTOR, '#default_0_vehicle_adj\.quoteandbind\.jppa\.directives\.JPPAQuoteDetailsCtrl\.Coverage\ for\ other\ person\ \&\ property_group > div:nth-child(1) > div > div > div > div.adjTariffPcBlock__price-body > div > div:nth-child(1) > gw-qnb-quote-view-cell > div > div.adjTariffPcBlock__price-body-inner.ng-scope').text
        compensation0=int(re.sub(r"\D", "", compensation0))
        compensation1=browser.find_element(By.CSS_SELECTOR, '#default_0_vehicle_adj\.quoteandbind\.jppa\.directives\.JPPAQuoteDetailsCtrl\.Coverage\ for\ other\ person\ \&\ property_group > div:nth-child(1) > div > div > div > div.adjTariffPcBlock__price-body > div > div:nth-child(2) > gw-qnb-quote-view-cell > div > div.adjTariffPcBlock__price-body-inner.ng-scope').text
        compensation1=int(re.sub(r"\D", "", compensation1))

        ###傷害
        injury0=browser.find_element(By.CSS_SELECTOR, '#default_1_vehicle_adj\.quoteandbind\.jppa\.directives\.JPPAQuoteDetailsCtrl\.Coverage\ for\ passengers_group > div:nth-child(1) > div > div > div > div.adjTariffPcBlock__price-body > div > div:nth-child(1) > gw-qnb-quote-view-cell > div > div.adjTariffPcBlock__price-body-inner.ng-scope').text
        injury0=int(re.sub(r"\D", "", injury0))
        injury1=browser.find_element(By.CSS_SELECTOR, '#default_1_vehicle_adj\.quoteandbind\.jppa\.directives\.JPPAQuoteDetailsCtrl\.Coverage\ for\ passengers_group > div:nth-child(1) > div > div > div > div.adjTariffPcBlock__price-body > div > div:nth-child(2) > gw-qnb-quote-view-cell > div > div.adjTariffPcBlock__price-body-inner.ng-scope > p').text
        injury1=int(re.sub(r"\D", "", injury1))

        ###車両
        physical0=browser.find_element(By.CSS_SELECTOR, '#default_2_vehicle_adj\.quoteandbind\.jppa\.directives\.JPPAQuoteDetailsCtrl\.Coverage\ for\ vehicle\ damage_group > div:nth-child(1) > div > div > div > div.adjTariffPcBlock__price-body > div > div:nth-child(1) > gw-qnb-quote-view-cell > div > div.adjTariffPcBlock__price-body-inner.ng-scope').text
        physical0=int(re.sub(r"\D", "", physical0))
        physical1=browser.find_element(By.CSS_SELECTOR, '#default_2_vehicle_adj\.quoteandbind\.jppa\.directives\.JPPAQuoteDetailsCtrl\.Coverage\ for\ vehicle\ damage_group > div:nth-child(1) > div > div > div > div.adjTariffPcBlock__price-body > div > div:nth-child(2) > gw-qnb-quote-view-cell > div > div.adjTariffPcBlock__price-body-inner.ng-scope').text
        physical1=int(re.sub(r"\D", "", physical1))

        ###その他
        other0=browser.find_element(By.CSS_SELECTOR, '#default_3_vehicle_adj\.quoteandbind\.jppa\.directives\.JPPAQuoteDetailsCtrl\.Other\ coverages_group > div:nth-child(1) > div > div > div > div.adjTariffPcBlock__price-body > div > div:nth-child(1) > gw-qnb-quote-view-cell > div > div.adjTariffPcBlock__price-body-inner.ng-scope').text
        other0=int(re.sub(r"\D", "", other0))
        other1=browser.find_element(By.CSS_SELECTOR, '#default_3_vehicle_adj\.quoteandbind\.jppa\.directives\.JPPAQuoteDetailsCtrl\.Other\ coverages_group > div:nth-child(1) > div > div > div > div.adjTariffPcBlock__price-body > div > div:nth-child(2) > gw-qnb-quote-view-cell > div > div.adjTariffPcBlock__price-body-inner.ng-scope').text
        other1=int(re.sub(r"\D", "", other1))


        data['車有P'] = result0_discount
        data['車無P'] = result1_discount
        data['イ割なし車有'] = result0
        data['イ割なし車無'] = result1
        data['車有P_賠償'] = compensation0
        data['車無P_賠償'] = compensation1
        data['車有P_傷害'] = injury0
        data['車無P_傷害'] = injury1
        data['車有P_車両'] = physical0
        data['車無P_車両'] = physical1
        data['車有P_その他'] = other0
        data['車無P_その他'] = other1


    #不測のエラーが起きた場合は、結果にEを入力する
    except :
        data['車有P']='E'
    
    browser.close()

    return data


# In[3]:



if __name__ == "__main__":
    FILE_NAME='AXA条件_データ1_定点'
    SHEET_NAME='AXA打鍵'

    #データ読み込み
    df=pd.read_excel(FILE_NAME+'.xlsm',sheet_name=SHEET_NAME)
    data=df.loc[84,:].to_dict()
    AXA_func(data) 
    print(data)

