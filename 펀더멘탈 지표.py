"""
네이버 증권 크롤링
"""
import pandas as pd
import requests
import numpy as np
from tqdm import tqdm

PATH_ = r"./JeWaJe_Excel.xlsx"

def crawl(code):
    URL = f"https://finance.naver.com/item/main.nhn?code={code}"
    r = requests.get(URL)
    tmp = pd.read_html(r.text)
    # 현재주가 얻기
    df = tmp[2]
    price = df.loc[1]['종가']
    price = np.array([price])

    # 재무정보 얻기
    df = tmp[3]
    df.set_index(df.columns[0],inplace=True)
    df.index.rename('주요재무정보', inplace=True)
    df.columns = df.columns.droplevel(2)
    annual_data = pd.DataFrame(df).xs('최근 연간 실적',axis=1)
    quater_data = pd.DataFrame(df).xs('최근 분기 실적',axis=1)

    # 연간 데이터만 우선 사용. 분기는 어떻게 쓸지 모르겠네...
    col_name = annual_data.columns.to_numpy()

    rev = annual_data.loc['매출액'].to_numpy()
    earn = annual_data.loc['영업이익'].to_numpy()
    earn_percent = annual_data.loc['영업이익률'].to_numpy()
    roe = annual_data.loc['ROE(지배주주)'].to_numpy()
    eps = annual_data.loc['EPS(원)'].to_numpy()
    bps = annual_data.loc['BPS(원)'].to_numpy()
    dps = annual_data.loc['주당배당금(원)'].to_numpy()

    annual = np.concatenate([price, rev, earn, earn_percent, roe, eps, bps, dps])
    for i in range(len(annual)):
        if annual[i] != '-' and annual[i] != 'nan':
            annual[i] = float(annual[i])

    return annual

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

def Fundamental():
    df1, df2, df3, df4 = OpenExcel()

    codes = df1["Code"].tolist()
    names = df1["이름"].tolist()

    fundamental = []

    for i in tqdm(range(len(codes)), desc="진행률"):
        code = str(codes[i])
        while len(code) < 6:
            code = "0" + code
        name = names[i]
        vals = [code, name]
        tmp = crawl(code)
        vals = np.append(vals, tmp)
        fundamental.append(vals)

    fundamental = np.array(fundamental)
    df4 = pd.DataFrame(fundamental, columns=['Code', '이름', '현재가',
                                             '매출19', '매출20', '매출21', '매출22',
                                             '영익19', '영익20', '영익21', '영익22',
                                             '이익률19', '이익율20', '이익율21', '이익율22',
                                             'roe19', 'roe20', 'roe21', 'roe22',
                                             'eps19', 'eps20', 'eps21', 'eps22',
                                             'bps19', 'bps20', 'bps21', 'bps22',
                                             'dps19', 'dps20', 'dps21', 'dps22'])

    SaveExcel(df1, df2, df3, df4)


Fundamental()