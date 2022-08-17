# -*- coding: utf-8 -*-
import random
import time

import pandas as pd
import schedule as schedule
from bs4 import BeautifulSoup
from nltk.tokenize import word_tokenize
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.jobstores.base import JobLookupError

import time
import EnterpriseCrawling
import WordCloud
import TopicModeling
import ReportDocx

#IBK, 하나, 국민, 농협, 신한, 우리, 카카오뱅크, 토스, 케이뱅크
url = ['https://play.google.com/store/apps/details?id=com.ibk.android.ionebank', 'https://play.google.com/store/apps/details?id=com.kebhana.hanapush',
       'https://play.google.com/store/apps/details?id=com.kbstar.kbbank', 'https://play.google.com/store/apps/details?id=nh.smart.banking',
       'https://play.google.com/store/apps/details?id=com.shinhan.sbanking', 'https://play.google.com/store/apps/details?id=com.wooribank.smart.npib',
       'https://play.google.com/store/apps/details?id=com.kakaobank.channel', 'https://play.google.com/store/apps/details?id=viva.republica.toss',
       'https://play.google.com/store/apps/details?id=com.kbankwith.smartbank']
csv = ['./reviews/개인고객/ibkbank_individual_review_dataset.csv', './reviews/개인고객/HANAreview_dataset.csv',
       './reviews/개인고객/KBreview_dataset.csv', './reviews/개인고객/NHreview_dataset.csv',
       './reviews/개인고객/신한review_dataset.csv', './reviews/개인고객/WONreview_dataset.csv',
       './reviews/인터넷뱅크/KAKAO_review_dataset.csv', './reviews/인터넷뱅크/TOSS_review_dataset.csv',
       './reviews/인터넷뱅크/KBank_review_dataset.csv']
total_rating = ['./reviews/별점/ibkbank_individual_review_rating.csv', './reviews/별점/HANAreview_rating.csv',
       './reviews/별점/KBreview_rating.csv', './reviews/별점/NHreview_rating.csv',
       './reviews/별점/신한review_rating.csv', './reviews/별점/WONreview_rating.csv',
       './reviews/별점/KAKAO_review_rating.csv', './reviews/별점/TOSS_review_rating.csv',
       './reviews/별점/KBank_review_rating.csv']

#개인앱 리뷰 크롤링
def main():
    #모든 은행 크롤링하기 위한 반복문
    for i in range(0, 9):
        chrome_driver = "chromedriver.exe"
        URL = url[i]
        CSV = csv[i]
        RATING = total_rating[i]

        # 크롬 드라이버 세팅
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--disable-blink-features=AutomationControlled')
        driver = webdriver.Chrome(chrome_driver, options=chrome_options)

        def scroll(modal):
            # 일정 개수만 크롤링 하기 위한 카운트
            count = 0
            try:
                while True:
                    count = count + 1
                    pause_time = random.uniform(0.5, 0.8)
                    # 최하단까지 스크롤
                    driver.execute_script("arguments[0].scrollTo(0, arguments[0].scrollHeight);", modal)
                    # 페이지 로딩 대기
                    time.sleep(pause_time)
                    # 무한 스크롤 동작을 위해 살짝 위로 스크롤
                    driver.execute_script("arguments[0].scrollTo(0, arguments[0].scrollHeight-50);", modal)
                    time.sleep(pause_time)
                    # 스크롤 높이 새롭게 받아오기
                    new_height = driver.execute_script("return arguments[0].scrollHeight", modal)
                    print(count)
                    try:
                        # '더보기' 버튼 있을 경우 클릭
                        all_review_button = driver.find_element_by_xpath(
                            '/html/body/div[1]/div[4]/c-wiz/div/div[2]/div/div/main/div/div[1]/div[2]/div[2]/div/span/span').click()
                    except:
                        # 스크롤 완료 경우
                        if count == 15:
                            print("스크롤 완료")
                            break

            except Exception as e:
                print("에러 발생: ", e)

        # 페이지 열기
        driver.get(URL)
        # 페이지 로딩 대기
        wait = WebDriverWait(driver, 5)

        # '리뷰 모두 보기' 버튼 렌더링 확인(path 수정 @2022-06-22)
        all_review_button_xpath = '/html/body/c-wiz[2]/div/div/div[1]/div[2]/div/div[1]/c-wiz[4]/section/div/div/div[5]/div/div/button/span'
        button_loading_wait = wait.until(EC.element_to_be_clickable((By.XPATH, all_review_button_xpath)))
        # '리뷰 모두 보기' 버튼 클릭
        # driver.find_element_by_xpath(all_review_button_xpath).click()
        driver.find_element(By.XPATH, all_review_button_xpath).click()  # 위에건 안되서 이렇게 수정함(셀레니움 버전 차이)

        # '리뷰 모두 보기' 페이지 렌더링 대기
        all_review_page_xpath = '/html/body/div[4]/div[2]/div/div/div/div/div[2]'
        page_loading_wait = wait.until(EC.element_to_be_clickable((By.XPATH, all_review_page_xpath)))

        # 페이지 무한 스크롤 다운
        modal = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, "//div[@class='fysCi']")))
        scroll(modal)

        # html parsing하기
        html_source = driver.page_source
        soup_source = BeautifulSoup(html_source, 'html.parser')

        driver.quit()

        # html 데이터 저장
        with open("./dataset/data_html.html", "w", encoding='utf-8') as file:
            file.write(str(soup_source))

        #앱 평점 데이터
        app_rating = soup_source.find_all(class_='jILTFe')[0].text
        #앱 평점 데이터 저장용 배열
        dataset_rating = []
        # 앱 평점 dataset_rating에 추가
        dataset_rating.append(app_rating)

        # 리뷰 데이터 클래스 접근
        review_source = soup_source.find_all(class_='RHo1pe')
        # 리뷰 데이터 저장용 배열
        dataset = []
        # 데이터 넘버링을 위한 변수
        review_num = 0
        
        # 리뷰 1개씩 접근해 정보 추출
        for review in review_source:
            review_num += 1
            # 리뷰 등록일 데이터 추출
            date_full = review.find_all(class_='bp9Aid')[0].text
            date_year = date_full[0:4]  # 연도 데이터 추출
            # 해당 단어가 등장한 인덱스 추출
            year_index = date_full.find('년')
            month_index = date_full.find('월')
            day_index = date_full.find('일')

            date_month = str(int(date_full[year_index + 1:month_index]))  # 월(Month) 데이터 추출
            # 월 정보가 1자리의 경우 앞에 0 붙이기(e.g., 1월 -> 01월)
            if len(date_month) == 1:
                date_month = '0' + date_month

            date_day = str(int(date_full[month_index + 1:day_index]))  # 일(Day) 데이터 추출
            # 일 정보가 1자리의 경우 앞에 0 붙여줌(e.g., 7일 -> 07일)
            if len(date_day) == 1:
                date_day = '0' + date_day

            # 리뷰 등록일 full version은 최종적으로 yyyymmdd 형태로 저장
            date_full = date_year + date_month + date_day
            user_name = review.find_all(class_='X5PpBb')[0].text  # 닉네임 데이터 추출
            rating = review.find_all(class_="iXRFPc")[0]['aria-label'][10]  # 평점 데이터 추출
            try:
                content = review.find_all(class_='h3YV2d')[0].text  # 리뷰 데이터 추출
            except IndexError:
                pass
            else:
                data = {
                    "id": review_num,
                    "date": date_full,
                    "dateYear": date_year,
                    "dateMonth": date_month,
                    "dateDay": date_day,
                    "rating": rating,
                    "userName": user_name,
                    "content": content
                }
                dataset.append(data)

        # 크롤링한 리뷰 csv 파일로 저장
        df = pd.DataFrame(dataset)
        df.to_csv(CSV, encoding='utf-8-sig')

        # 저장한 리뷰 정보 불러오기
        df = pd.read_csv(CSV, encoding='utf-8-sig')
        df = df.drop(['Unnamed: 0'], axis=1)  # 불필요한 칼럼 삭제

        # 앱 평정 csv 파일로 저장
        df_rating = pd.DataFrame(dataset_rating)
        df_rating.to_csv(RATING, encoding='utf-8-sig')

        # 저장한 평점 정보 불러오기
        df_rating = pd.read_csv(RATING, encoding='utf-8-sig')
        df_rating = df_rating.drop(['Unnamed: 0'], axis=1)  # 불필요한 칼럼 삭제



    # 기업 크롤링 코드 연결
    EnterpriseCrawling.main()

    # 워드클라우드 코드 연결
    WordCloud.main()

    # 토픽모델링 코드 연결 (오래걸려서 실행X, csv파일 추출해둠)
    # TopicModeling.main()

    # 보고서 코드 연결
    ReportDocx.main()
        

# 전체 메인함수
main()

#스케줄링
# sched = BackgroundScheduler()
# sched.start()
# #매월 1일 오전 9시에 main 함수 실행
# sched.add_job(main, 'cron', day=1, hour=9)
#
# while True:
#     time.sleep(1)
