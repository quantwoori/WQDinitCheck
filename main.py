from func import *
from typing import List, Dict
import shutil


def step0():
    print("+----------------------+")
    print("| Stock Adding Process |")
    print("+----------------------+")

    print("본 프로그램은 새로운 종목을 대차잔고 모니터링 모델에 추가해 줄것 입니다.")
    print("각 단계가 끝나면, 엔터(Enter)키를 눌러주시면 다음단계로 진행됩니다.")
    print(">>>엔터를 눌러주세요")

    input()


def step1() -> List:
    print("\n\n")
    print("===========1 단 계===========")
    print("1단계. 추가하고 싶으신 종목들의 티커(6자리 종목 코드)를 다음과 같이 넣어주세요")
    print("종목들을 모두 치셨다면 엔터를 눌러주세요")
    print("ex) 삼성전자와 하이닉스를 추가하고 싶으실 때")
    print("005930, 000660")

    while True:
        try:
            print("1단계. 추가하고 싶으신 종목들의 티커(6자리 종목 코드)를 다음과 같이 넣어주세요")
            print("종목들을 모두 치셨다면 엔터를 눌러주세요")
            print("ex) 삼성전자와 하이닉스를 추가하고 싶으실 때")
            print("005930, 000660")

            stks = list(input().replace(' ', '').split(','))

            print("종목 코드가 맞는지 확인 중입니다.")
            check_stks = all(map(lambda x: len(x) == 6, stks))
            if not check_stks:
                raise Exception
        except Exception:
            print("종목 코드를 잘못 입력하셨습니다.")
            print(f"입력값: {stks}")
            print("해당 단계를 처음 부터 시작해주세요")
            print("\n\n")
        else:
            print("종목 코드 확인되었습니다.")
            print(f"입력값: {stks}")
            print("1단계 정상종료")

            # Insert excel function in new.xlsx
            cd = CheckData()
            s, e = (
                (datetime.today() - timedelta(days=40)).strftime("%Y%m%d"),
                (datetime.today() - timedelta(days=1)).strftime("%Y%m%d")
            )
            cd.xl_cell_input(
                start_date=s,
                end_date=e,
                stk_codes=stks
            )
            return stks


def step2():
    print("\n\n")
    print("==========2 단 계==========")
    print("2단계. 이 단계는 추가하실 종목데이터를 체크 함수로부터 가져오는 역할을 수행합니다.")
    print("new.xlsx파일에 있는 자료를 모두 newinsert.xlsx파일로 옮겨주세요")
    print("옮기실 때는 값복사를 하셔야 정상적으로 수행됩니다.")
    print("값복사를 완료하시고, 엑셀을 저장해주세요.")
    print(">>>엔터를 눌러주세요")
    input()


def step3(stock_list:list) -> (XLClean, Dict):
    print("\n\n")
    print("==========3 단 계==========")
    print("3단계. 이 단계에선 프로그램이 불러주는 정보를 Sstkins.py에 수동으로 기록하셔야 합니다.")
    print("먼저 Sstkins.py 파일(혹은 바로가기 파일)을 >> 오른쪽클릭 >> 연결 프로그램 >> 메모장")
    print("메모장으로 열어주세요")
    print("\n")
    print("하단에 기록된 방법대로 지금부터 프로그램이 불러주는 정보를 입력하시면 됩니다.")
    x = XLClean("newinsert.xlsx")
    req_param = 2
    stks = [x.data.columns[i] for i in range(len(x.data.columns))
            if (i - req_param + 1) % 4 == 0]

    updater_temp = dict()
    for k, v in zip(stks, stock_list):
        print("정확하게 다음과 같이 입력하세요. 입력하신 후 엔터를 누르면 다음 것으로 넘어갑니다")
        print(f'"{k}": "{v}"')
        updater_temp[k] = v
        input()
    print(f"{stock_list}에 들어있는 데이터 입력이 완료되었습니다.")
    print()

    print("입력한 메모장을 저장하신 다음 종료하세요.")
    print(">>>엔터를 눌러주세요")
    input()

    return x, updater_temp



def step4(cleaner:XLClean, new:dict):
    print("\n\n")
    print("==========4 단 계==========")
    print("4단계. 이 단계는 자동으로 수행됩니다. 이 단계는 값복사하신 데이터를 데이터베이스에 넣는 작업을 수행합니다.")
    print("만약 오류가 나고 꺼졌다면, 그 말은 2, 3단계에 문제가 있다는 것을 의미합니다.")
    print("오류가 뜬다면 작업 내역과 데이터를 확인하세요")
    # 한번 파일이 실행된 상태에서는 락된 상태로 들어가기 때문에, 아무것도 안됨
    # 지금은 new라는 것을 이용해서 업데이트 한다.
    cleaner.kospi.k100['STK_CODE'].update(new)

    d = cleaner.clean_column(cleaner.data, 2)
    d.date = d.date.apply(cleaner.clean_date)
    d.stock = d.stock.apply(cleaner.clean_stock)

    # Database
    DB_COL = ['date', 'borrow', 'd_borrow', 'r_borrow', 'stkcode']
    d_db = d.to_numpy().tolist()
    d_db = [tuple(_) for _ in d_db]
    cleaner.server.insert_row(
        table_name='RAWborrow',
        schema='dbo',
        database='WSOL',
        col_=DB_COL,
        rows_=d_db
    )
    print("로데이터(Raw Data)가 데이터베이스에 삽입되었습니다.")

def hiddenstep(from_loc:str, to_loc:str):
    shutil.copy(from_loc, to_loc)


if __name__ == "__main__":
    step0()
    s1r = step1()
    step2()
    s2r, s2d = step3(s1r)
    step4(s2r, s2d)
    #
    # loc0, loc1 = (
    #     "C:/Users/check/move/initCheck/settings/Sstkins.py",
    #     "C:/Users/check/move/checkData/settings/Sstk.py"
    # )
    # hiddenstep(loc0, loc1)
    #
    #
    #
    #
