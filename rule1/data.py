import numpy as np

from enum import Enum
from urllib.parse import quote
from exceptions import FieldNotFoundException


class Section(Enum):
    INCOME = quote('Income Statement')
    CASH_FLOW = quote('Cash Flow')
    BALANCE_SHEET = quote('Balance Sheet')
    METRICS = quote('Metrics')
    GROWTH = quote('Growth')


class Field(Enum):
    SALES_GROWTH = 'Sales_Growth'
    ROIC = 'ROIC'
    EPS_GROWTH = 'EPS_Growth'
    FCF_GROWTH = 'FCF_Growth'
    BVPS_GROWTH = 'BVPS_Growth'


class Indicator:
    def __init__(self, name: Field, values):
        numbers = [x for x in values if x is not None]

        self.name = name.value
        self.year10 = np.average(numbers[-10:])
        self.year5 = np.average(numbers[-5:])
        self.year1 = np.average(numbers[-1:])


class CompanyData:
    lookup_table = {
        Field.SALES_GROWTH: (Section.INCOME, 'Revenue Growth'),
        Field.ROIC: (Section.METRICS, 'ROIC'),
        Field.EPS_GROWTH: (Section.GROWTH, 'EPS Growth'),
        Field.FCF_GROWTH: (Section.GROWTH, 'Free Cash Flow growth'),
        Field.BVPS_GROWTH: (Section.GROWTH, 'Book Value per Share Growth'),
    }

    def __init__(self, ticker, income_statement, balance_sheet, cash_flow, metrics, growth):
        self.ticker = ticker
        self.data = {
            Section.INCOME: income_statement,
            Section.BALANCE_SHEET: balance_sheet,
            Section.CASH_FLOW: cash_flow,
            Section.METRICS: metrics,
            Section.GROWTH: growth,
        }

        self.income_statement = income_statement
        self.balance_sheet = balance_sheet
        self.cash_flow = cash_flow
        self.metrics = metrics
        self.growth = growth

    def lookup_row(self, field: Field):
        section, title = CompanyData.lookup_table[field]
        data = self.data[section]
        title_column = [data.cell(row=r, column=1).value for r in range(1, data.max_row + 1)]
        matching_cell = [idx for idx, cell in enumerate(title_column) if cell == title]

        if not matching_cell:
            raise FieldNotFoundException(section, title)

        values = [data.cell(row=matching_cell[0] + 1, column=c).value for c in range(2, max(data.max_column + 1, 12))]

        return values

    def get_indicator(self, field: Field):
        values = self.lookup_row(field)
        indicator = Indicator(field, values)

        return indicator

    def get_roic(self):
        return self.get_indicator(Field.ROIC)

    def get_eps_growth(self):
        return self.get_indicator(Field.EPS_GROWTH)

    def get_equity_growth(self):
        return self.get_indicator(Field.BVPS_GROWTH)

    def get_fcf_growth(self):
        return self.get_indicator(Field.FCF_GROWTH)

    def get_sales_growth(self):
        return self.get_indicator(Field.SALES_GROWTH)
