"""
엑셀 파일에 자동으로 데이터 채우는 프로그램
https://wikidocs.net/5756
책 참고
"""
import os
import sys
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
import pandas as pd
import numpy as np
import time
import datetime
from tqdm import tqdm

TR_REQ_TIME_INTERVAL = 0.2

Today = str(datetime.datetime.now())
Today = Today[0:4] + Today[5:7] + Today[8:10]

PATH_ = r"./JeWaJe_Excel.xlsx"

# 키움증권 API
class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()
        self._create_kiwoom_instance()
        self._set_signal_slots()
        self.price_data = []

    def _create_kiwoom_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")

    def _set_signal_slots(self):
        self.OnEventConnect.connect(self._event_connect)
        self.OnReceiveTrData.connect(self._receive_tr_data)

    def comm_connect(self):
        self.dynamicCall("CommConnect()")
        self.login_event_loop = QEventLoop()
        self.login_event_loop.exec_()

    def _event_connect(self, err_code):
        if err_code == 0:
            print("connected")
        else:
            print("disconnected")

        self.login_event_loop.exit()

    def get_code_list_by_market(self, market):
        code_list = self.dynamicCall("GetCodeListByMarket(QString)", market)
        code_list = code_list.split(';')
        return code_list[:-1]

    def get_master_code_name(self, code):
        code_name = self.dynamicCall("GetMasterCodeName(QString)", code)
        return code_name

    def set_input_value(self, id, value):
        self.dynamicCall("SetInputValue(QString, QString)", id, value)

    def comm_rq_data(self, rqname, trcode, next, screen_no):
        self.dynamicCall("CommRqData(QString, QString, int, QString", rqname, trcode, next, screen_no)
        self.tr_event_loop = QEventLoop()
        self.tr_event_loop.exec_()

    def _comm_get_data(self, code, real_type, field_name, index, item_name):
        ret = self.dynamicCall("CommGetData(QString, QString, QString, int, QString", code,
                               real_type, field_name, index, item_name)
        return ret.strip()

    def _get_repeat_cnt(self, trcode, rqname):
        ret = self.dynamicCall("GetRepeatCnt(QString, QString)", trcode, rqname)
        return ret

    def _receive_tr_data(self, screen_no, rqname, trcode, record_name, next, unused1, unused2, unused3, unused4):
        if next == '2':
            self.remained_data = True
        else:
            self.remained_data = False

        if rqname == "opt10081_req":
            self._opt10081(rqname, trcode)

        if rqname == "opt10015_req":
            self._opt10015(rqname, trcode)

        try:
            self.tr_event_loop.exit()
        except AttributeError:
            pass

    def _opt10081(self, rqname, trcode):
        # 데이터 개수 가져오기
        data_cnt = self._get_repeat_cnt(trcode, rqname)

        for i in range(data_cnt):
            date = self._comm_get_data(trcode, "", rqname, i, "일자")
            #open = self._comm_get_data(trcode, "", rqname, i, "시가")
            #high = self._comm_get_data(trcode, "", rqname, i, "고가")
            #low = self._comm_get_data(trcode, "", rqname, i, "저가")
            close = self._comm_get_data(trcode, "", rqname, i, "현재가")
            #volume = self._comm_get_data(trcode, "", rqname, i, "거래량")
            self.price_data.append([date, int(close)])

    def _opt10015(self, rqname, trcode):
        date = self._comm_get_data(trcode, "", rqname, 0, "일자")
        price = self._comm_get_data(trcode, "", rqname, 0, "종가")
        print(date, price)

# 종목 코드 채우는 코드
# 나중에 코드 수정 -- NaN이 있을 때 거기만 채우는 방식으로 바꿔야할것 같다
# 이건 사실 잘 안써서 안건드리긴 함
def NewDataAdd():
    app = QApplication(sys.argv)
    kiwoom = Kiwoom()
    kiwoom.comm_connect()

    stocks = dict()

    kosdaq_code = kiwoom.get_code_list_by_market('10')
    kospi_code = kiwoom.get_code_list_by_market('0')

    for code in kospi_code:
        name = kiwoom.get_master_code_name(code)
        stocks[name] = code

    for code in kosdaq_code:
        name = kiwoom.get_master_code_name(code)
        stocks[name] = code

    df = pd.read_excel("./JeWaJe_Excel.xlsx", sheet_name="유니버스")

    names = df["이름"].tolist()

    codes = []
    for n in names:
        c = stocks[n]
        codes.append(c)

    df = df.to_numpy()
    df = np.delete(df, 0, axis=1)
    new = []

    for i, s in enumerate(df):
        new.append(np.insert(s, 0, codes[i]))

    new = np.array(new)
    new_df = pd.DataFrame(new, columns=["Code", "이름", "섹터", "기타"])
    os.remove("./JeWaJe_Excel.xlsx")
    new_df.to_excel("JeWaJe_Excel.xlsx", sheet_name="유니버스", index=False)

# 종목 코드 입력하면 종목 모멘텀 값 리턴
def SeveralDays(kiwoom, code, date=Today):
    # 데이터 초기화
    kiwoom.price_data = []
    # opt10081 TR 요청
    kiwoom.set_input_value("종목코드", code)
    kiwoom.set_input_value("기준일자", date)
    kiwoom.set_input_value("수정주가구분", 1)
    kiwoom.comm_rq_data("opt10081_req", "opt10081", 0, "0101")
    price = kiwoom.price_data

    now_price = price[0][1]

    five_day_mean = (price[0][1] + price[1][1] + price[2][1] + price[3][1] + price[4][1]) / 5
    five_day_ago = price[4][1]

    flag1 = False
    flag2 = False

    four_weak_mean = 0
    if (len(price) >= 20):
        for i in range(20):
            four_weak_mean += price[i][1]
        four_weak_mean = four_weak_mean / 20
        four_weak_ago = price[19][1]
    else:
        four_weak_mean = now_price
        four_weak_ago = now_price
        flag1 = True

    three_month_mean = 0
    if (len(price) >= 60):
        for i in range(60):
            three_month_mean += price[i][1]
        three_month_mean = three_month_mean / 60
        three_month_ago = price[59][1]
    else:
        three_month_mean = now_price
        three_month_ago = now_price
        flag2 = True

    momentum = np.array([five_day_ago, five_day_mean, four_weak_ago, four_weak_mean, three_month_ago, three_month_mean])
    momentum = ((now_price - momentum) / momentum) * 100
    momentum = np.round(momentum, 2)

    if flag1:
        momentum[2] = None
        momentum[3] = None
    if flag2:
        momentum[4] = None
        momentum[5] = None

    return momentum

# DataFrame
def MomentumFill():
    app = QApplication(sys.argv)
    kiwoom = Kiwoom()
    kiwoom.comm_connect()

    df1, df2, df3, df4 = OpenExcel()

    codes = df1["Code"].tolist()
    names = df1["이름"].tolist()

    momentum = []

    for i in tqdm(range(len(codes)), desc="진행률"):
        if (i % 50 == 0) and (i != 0):
            time.sleep(30)
        code = str(codes[i])
        while len(code) < 6:
            code = "0" + code
        vals = [code, names[i]]
        tmp = SeveralDays(kiwoom, code)
        vals = np.append(vals, tmp)
        momentum.append(vals)
        time.sleep(TR_REQ_TIME_INTERVAL)

    momentum = np.array(momentum)
    df3 = pd.DataFrame(momentum, columns=["Code", "이름", "5일전", "5일평균", "20일전", "20일평균", "60일전", "60일평균"])
    SaveExcel(df1, df2, df3, df4)

# 판다스로 엑셀시트 열때 모든 시트 다 불러오기
def OpenExcel(path=PATH_):
    df1 = pd.read_excel(path, sheet_name="총점")
    df2 = pd.read_excel(path, sheet_name="유니버스")
    df3 = pd.read_excel(path, sheet_name="모멘텀 지표")
    df4 = pd.read_excel(path, sheet_name="펀더멘탈 지표")
    return(df1, df2, df3, df4)

def SaveExcel(df1, df2, df3, df4, path=PATH_):
    writer = pd.ExcelWriter(PATH_, engine="xlsxwriter")
    df1.to_excel(writer, sheet_name="총점", index=False)
    df2.to_excel(writer, sheet_name="유니버스", index=False)
    df3.to_excel(writer, sheet_name="모멘텀 지표", index=False)
    df4.to_excel(writer, sheet_name="펀더멘탈 지표", index=False)
    writer.save()

    writer = pd.ExcelWriter(r"./JeWaJe_Excel_copy.xlsx", engine="xlsxwriter")
    df1.to_excel(writer, sheet_name="총점", index=False)
    df2.to_excel(writer, sheet_name="유니버스", index=False)
    df3.to_excel(writer, sheet_name="모멘텀 지표", index=False)
    df4.to_excel(writer, sheet_name="펀더멘탈 지표", index=False)
    writer.save()

MomentumFill()
