import openpyxl
import aiohttp
import io

from tqdm import tqdm
from os import path, makedirs
from datetime import datetime
from data import Section, CompanyData


def get_cache_file_name(ticker, section):
    now = datetime.now()
    q_stamp = f'{now.year}_Q{1 + int(now.month / 4)}'
    file_name = f'{ticker}_{str(section.name)}_{q_stamp}.xlsx'

    return file_name


def save_statement_to_cache(ticker: str, section: Section, content: bytes):
    home_path = path.expanduser("~")
    cache_path = path.join(home_path, '.local', 'share', 'rule1', 'cache')

    if not path.isdir(cache_path):
        makedirs(cache_path)

    statement_path = path.join(cache_path, get_cache_file_name(ticker, section))

    with open(statement_path, 'wb') as cache_file:
        cache_file.write(content)


def check_statement_in_cache(ticker: str, section: Section):
    home_path = path.expanduser("~")
    cache_file_path = path.join(home_path, '.local', 'share', 'rule1', 'cache', get_cache_file_name(ticker, section))

    if path.isfile(cache_file_path):
        workbook = openpyxl.load_workbook(cache_file_path)

        return workbook

    return None


def get_company_url(ticker, section: Section):
    scheme = 'https'
    host = 'stockrow.com'
    path = f'/api/companies/{ticker}/financials.xlsx'
    query = f'dimension=MRY&section={section.value}&sort=asc'

    return f'{scheme}://{host}{path}?{query}'


async def get_statement_worksheet(ticker, section: Section):
    cached_statement = check_statement_in_cache(ticker, section)

    if cached_statement is not None:
        return cached_statement[ticker]

    url = get_company_url(ticker, section)

    async with aiohttp.ClientSession() as session:
        async with session.get(url) as response:
            content = await response.content.read()
            xlsx = io.BytesIO(content)
            data = openpyxl.load_workbook(xlsx)

            save_statement_to_cache(ticker, section, content)

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
    income, balance_sheet, cash_flow, metrics, growth = [await t for t in tqdm(tasks, desc=description)]
    company_data = CompanyData(ticker, income, balance_sheet, cash_flow, metrics, growth)

    return company_data


async def lookup_companies(tickers):
    companies = {}

    for ticker in tickers:
        company_data = await get_company_data(ticker)
        companies[ticker] = company_data

    return companies
