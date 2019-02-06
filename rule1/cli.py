import argparse
import re
import os.path


class InvalidInputException:
    def __init__(self, content):
        self.content = content


def extract_tickers(content: str):
    trimmed_content = content.replace(' ', '')
    if not re.match('^[A-Za-z0-9,]*$', trimmed_content):
        raise InvalidInputException(trimmed_content)

    tickers = trimmed_content.split(',')

    return tickers


def parse_arguments():
    parser = argparse.ArgumentParser(description='Simple command line utility to check health indicators and calculate '
                                                 'MOS price for companies. Based on Phil Town\'s Rule #1 investing.')
    input_group = parser.add_mutually_exclusive_group()
    input_group.add_argument('-t', '--tickers', type=str, metavar='', help='Comma separated list of stock tickers')
    input_group.add_argument('-i', '--input', type=str, metavar='', help='Path to an input file containing '
                                                                         'a comma separated list of stock tickers')
    parser.add_argument('-o', '--output', type=str, metavar='', required=True, help='Path to the output file')
    args = parser.parse_args()

    if args.input:
        with open(args.input) as input_file:
            ticker_list = input_file.read()
    else:
        ticker_list = args.tickers

    tickers = extract_tickers(ticker_list)
    output_path = args.output

    return tickers, output_path
