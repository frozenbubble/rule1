import asyncio

from cli import parse_arguments
from stockrow import lookup_companies
from reports import generate_reports


async def main():
    tickers, output_path = parse_arguments()
    companies = await lookup_companies(tickers)
    generate_reports(output_path, companies)


loop = asyncio.get_event_loop()
loop.run_until_complete(main())
