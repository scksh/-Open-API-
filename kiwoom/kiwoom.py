from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from config.errorCode import *
from datetime import datetime, timedelta
#import requests  # 여기서 requests 모듈을 임포트합니다.
import pandas as pd
today = datetime.today().date()
today_str = today.strftime("%Y%m%d")


class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()

        print("Kiwoom 클래스")

        ###### eventloop ########
        self.login_event_loop = None
        self.detail_account_info_event_loop = None
        #########################

        ###### variables ########
        self.account_num = None
        self.code = "005930"
        self.daily_chart = None
        #########################

        ######종목분석용########
        self.calcul_data = []
        #####################

        self.get_ocx_instance()
        self.event_slots()

        self.signal_login_commConnect()
        self.get_account_info()
        self.detail_account_info()
        self.get_daily_chart()


    def get_ocx_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")

    def event_slots(self):
        self.OnEventConnect.connect(self.login_slot)
        self.OnReceiveTrData.connect(self.trdata_slot)

    def login_slot(self, errCode):
        print(errCode)
        print(errors(errCode))


        self.login_event_loop.exit()


    def signal_login_commConnect(self):
        self.dynamicCall("CommConnect()")

        self.login_event_loop = QEventLoop()
        self.login_event_loop.exec_()

    def get_account_info(self):
        account_list = self.dynamicCall("GetLoginInfo(string)", "ACCNO")

        self.account_num = account_list.split(";")[0]
        print(f"내 보유 계좌번호 {self.account_num}")  # 8089266711

    def detail_account_info(self):
        print("예수금을 요청하는 부분")

        self.dynamicCall("SetInputValue(String, String)", "계좌번호", "8089266711")
        self.dynamicCall("SetInputValue(String, String)", "비밀번호", "0000")
        self.dynamicCall("SetInputValue(String, String)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(String, String)", "조회구분", "2")
        self.dynamicCall("CommRqData(string, String, int, String)", "예수금상세현황요청", "opw00001", "0", "2000")

        self.detail_account_info_event_loop = QEventLoop()
        self.detail_account_info_event_loop.exec_()

    def trdata_slot(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext):
        print("trdata_slot 호출됨")  # 디버깅용
        '''tr요청을 받는 슬롯
        sScrNo : 스크린번호
        sRQName : 요청할때 정한 이름 
        sTrCode : tr코드
        sRecordName : 사용안함
        sPrevNext : NextPage 유무

        return: 
        '''

        if sRQName == "예수금상세현황요청":
            deposit = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "예수금")
            print(f"예수금: {deposit}")
            print("예수금 형변환 %s" % int(deposit))

            rt_deposit = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "출금가능금액")
            print(f"출금가능금액: {rt_deposit}")
            print("출금가능금액 형변환 %s" % int(rt_deposit))

            self.detail_account_info_event_loop.exit()
        print("trdata_slot 완료")  # 추가

        if sRQName == "주식일봉차트조회요청":
            cnt = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)
            for i in range(cnt):  # 2*cnt가 아니라 cnt까지만 반복
                data = []
                current_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                                 "현재가")
                date = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "일자")
                start_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "시가")
                high_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "고가")
                low_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "저가")
                volume = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "거래량")

                data.append(date.strip())
                data.append(start_price.strip())
                data.append(current_price.strip())
                data.append(high_price.strip())
                data.append(low_price.strip())
                data.append(volume.strip())
                self.calcul_data.append(data.copy())

            # 다음 페이지가 있고, 데이터가 1200개 미만일 때 연속 조회를 요청
            if len(self.calcul_data) < 1200 and sPrevNext == "2":
                last_date = self.calcul_data[-1][1]  # 마지막 날짜로 갱신
                self.get_daily_chart(code=self.code, date=last_date, sPrevNext="2")
            else:
                # 데이터 수집 완료
                self.daily_chart = self.calcul_data.copy()
                self.daily_chart.reverse()
                print("1200일치 데이터 수집 완료 (역순 정렬됨)")
                print("데이터 개수:", len(self.daily_chart))  # 데이터 개수 출력
                print(self.daily_chart[:5])  # 첫 5개 데이터 출력 (예시)

                try:
                    self.save_data_to_csv(self.daily_chart)
                    print("데이터가 CSV 파일로 성공적으로 저장되었습니다.")
                except Exception as e:
                    print("CSV 저장 중 오류 발생:", e)


    def get_daily_chart(self, code=None, date=None, sPrevNext="0"):
        self.dynamicCall("SetInputValue(String, String)", "종목코드", code or self.code)
        self.dynamicCall("SetInputValue(String, String)", "수정주가구분", "1")
        self.dynamicCall("SetInputValue(String, String)", "기준일자", date or today_str)
        self.dynamicCall("CommRqData(string, String, int, String)", "주식일봉차트조회요청", "opt10081", sPrevNext, "1000")



    def save_data_to_csv(self, daily_data, filename='daily_chart.csv'):
        # 데이터프레임으로 변환
        df = pd.DataFrame(daily_data, columns=['Date', 'Start Price', 'Current Price', 'High Price', 'Low Price', 'Volume'])
        # CSV 파일로 저장
        df.to_csv(filename, index=False)
        print(f"데이터가 {filename}로 저장되었습니다.")



        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        