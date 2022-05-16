from dbms.DBmssql import MSSQL
from settings.Sstkins import KOSPI

import openpyxl as op
import pandas as pd

from datetime import datetime, timedelta


class XLClean:
    def __init__(self, excel_file:str):
        self.data = pd.read_excel(excel_file)
        self.server = MSSQL.instance()
        self.server.login(
            id='wsol1',
            pw='wsol1'
        )
        self.kospi = KOSPI()

    @staticmethod
    def clean_date(excel_date:int, dfmt:str="%Y%m%d") -> str:
        dt = datetime.fromordinal(
            int(
                datetime(1900, 1, 1).toordinal() +
                excel_date - 2
            )
        )
        return dt.strftime(dfmt)

    def clean_stock(self, stock_name:str):
        return self.kospi.k100['STK_CODE'][stock_name]

    @staticmethod
    def clean_column(df:pd.DataFrame, num_col:int) -> pd.DataFrame:
        """
        CLEAN CHECK DATA MESS INTO SOMETHING USABLE
        :param num_col:
        How many parameters did you requested?
        :return:
        pd.DataFrame
        """
        GROUP_COL = 2 + num_col
        SEG_COL = ['date', 'borrowed', 'd_borrowed', 'r_borrowed']

        result = pd.DataFrame(None)
        for i in range(0, len(df.columns), GROUP_COL):
            name = df.columns[i : (i + GROUP_COL)][1]
            seg = df[df.columns[i : (i + GROUP_COL)]][1:]  # First Column is a Bogus
            seg.columns = SEG_COL
            seg['stock'] = [name] * len(seg)

            result = pd.concat([result, seg])
        result = result.reset_index(drop=True)
        result = result.dropna()
        return result


class CheckData:
    def __init__(self, path:str="new.xlsx", result_path:str='result.xlsx'):
        # CONSTANT
        self.COST = 4
        self.ROW_START, self.COL_START = 1, 1

        # Excel Workfile
        self.wf = op.Workbook()
        self.dpath = path
        self.rpath = result_path

        # Target Stocks
        self.kospi = KOSPI()

    def xl_func_writer(self, start_date:str, end_date:str, stock_code:str) -> str:
        func = f'=CH("{start_date}", "{end_date}", -1, "D", FALSE, "{stock_code}", "14212,14214,14216", "ASC", "withtable=true;")'
        return func

    def xl_cell_input(self, start_date:str, end_date:str, stk_codes:set) -> None:
        ws = self.wf.active
        r = 0
        for s in stk_codes:
            val = self.xl_func_writer(start_date, end_date, stock_code=s)
            ws.cell(row=self.ROW_START,
                    column=self.COL_START + r * self.COST).value = val
            r += 1

        self.wf.save(self.dpath)

    def process_rpa_res(self, loc=r'C:\Users\Check\Documents\Quant_모니터링\result.xlsx'):
        d = pd.read_excel(loc)
