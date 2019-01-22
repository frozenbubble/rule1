from enum import Enum
from urllib.parse import quote
import openpyxl
import numpy as np
import aiohttp
import asyncio
import io
from tqdm import tqdm


class Section(Enum):
    INCOME = quote('Income Statement')
    CASH_FLOW = quote('Cash Flow')
    BALANCE_SHEET = quote('Balance Sheet')
    METRICS = quote('Metrics')
    GROWTH = quote('Growth')


class CompanyData:
    lookup_table = {
        'ROIC': (Section.METRICS, 41, 'ROIC'),
        'EPS_Growth': (Section.GROWTH, 6, 'EPS Growth'),
        'FCF_Growth': (Section.GROWTH, 11, 'Free Cash Flow growth'),
        'BVPS_Growth': (Section.GROWTH, 28, 'Book Value per Share Growth'),
        'Sales_Growth': (Section.INCOME, 3, 'Revenue Growth')
    }

    def __init__(self, income_statement, balance_sheet, cash_flow, metrics, growth):
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

    def get_roic(self):
        section, row, title = CompanyData.lookup_table['ROIC']
        data = self.data[section]
        title_cell = f'A{row}'

        if data[title_cell].value != title:
            print('Schema mismatch!')
            exit()

        roic_values = [data.cell(row=row, column=2 + i).value for i in range(10)]

        return np.average(roic_values)


def get_company_url(ticker, section: Section):
    scheme = 'https'
    host = 'stockrow.com'
    path = f'/api/companies/{ticker}/financials.xlsx'
    query = f'dimension=MRY&section={section.value}&sort=asc'

    return f'{scheme}://{host}{path}?{query}'


async def get_statement_worksheet(ticker, section: Section):
    url = get_company_url(ticker, section)

    async with aiohttp.ClientSession() as session:
        async with session.get(url) as response:
            content = await response.content.read()
            xlsx = io.BytesIO(content)
            data = openpyxl.load_workbook(xlsx)

            return data[ticker]


async def get_company_data(ticker):
    tasks = [
        get_statement_worksheet(ticker, Section.INCOME),
        get_statement_worksheet(ticker, Section.BALANCE_SHEET),
        get_statement_worksheet(ticker, Section.CASH_FLOW),
        get_statement_worksheet(ticker, Section.METRICS),
        get_statement_worksheet(ticker, Section.GROWTH)
    ]

    description = f'Downloading financial data for {ticker}'
    income, balance_sheet, cash_flow, metrics, growth = [await t for t in tqdm(tasks, desc=description)] #await asyncio.gather(*tasks)

    return CompanyData(income, balance_sheet, cash_flow, metrics, growth)


async def lookup_companies(tickers):
    companies = {}

    for ticker in tickers:
        company_data = await get_company_data(ticker)
        companies[ticker] = company_data

    return companies


async def main():
    companies = await lookup_companies(['TSLA', 'AAPL', 'AMZN', 'MSFT'])
    # aapl = companies['aapl']
    # print(aapl.get_roic())
    # tsla = await get_company_data('TSLA')

loop = asyncio.get_event_loop()
loop.run_until_complete(main())
