import os
import datetime

import pandas as pd
import packaging.version

import tqdm
import yfinance as yf
from scrape import Scraper

from scoring import *
from utils import *
from constants import *


# TODO - Write Dividend Yield Channels
# TODO - Draw DYC chart

def download_unscraped_data(lst):
    scraper = Scraper()
    scraper.login()
    unscraped = lst[lst.Ticker.apply(is_data_scraped) == False].Ticker.tolist()
    for ticker in unscraped:
        scraper.download_historical_data(ticker)


def download_current_data_for_all(tickers):
    if not isinstance([], list):
        raise TypeError('Wrong parameter type')

    if not isinstance(tickers[0], str):
        raise TypeError('Ticker list must contain strings')

    df = pd.DataFrame(columns=[NAME, TICKER, CURRENT_PRICE,
                               CURRENT_DPS, CURRENT_YLD, CREDIT_RATING, BETA, SECTOR])
    data = yf.Tickers(' '.join(tickers))
    for ticker in tickers:
        print(
            f'--------------- Downloading current data for {ticker} ---------------')
        tkr = data.tickers[ticker]
        name = tkr.info['longName']

        try:
            dps = tkr.info['dividendRate'] / 4
            price = tkr.history(period="1d").Close[0]
            yld = dps * 4 / price
            beta = tkr.info['beta']
            sector = tkr.info['sector']
        except (IndexError, KeyError) as exc:
            print(f'"No data available for {ticker}: {str(exc)}, skipping')
            continue

        df = df.append({
            NAME: name,
            TICKER: ticker,
            CURRENT_PRICE: price,
            CURRENT_DPS: dps,
            CURRENT_YLD: yld,
            CREDIT_RATING: 'AAA',
            BETA: beta,
            SECTOR: sector
        }, ignore_index=True)
    return df


def get_format_module():
    version = packaging.version.parse(pd.__version__)
    if version < packaging.version.parse('0.18'):
        return pd.core.format
    elif version < packaging.version.parse('0.20'):
        return pd.formats.format
    elif version < packaging.version.parse('0.24'):
        return pd.io.formats.excel
    else:
        return pd.io.formats.excel.ExcelFormatter


scraper = Scraper()
# scraper.login()
y_sp = 1.64

ccc = get_ccc()
file_path = os.path.join(os.getcwd(), "company-list.xlsx")
lst = pd.read_excel(file_path, header=6).iloc[:, 1:]
unanalyzed = lst.loc[lst.DCF.isnull(), 'Ticker'].tolist()
unscraped = lst[lst.Ticker.apply(is_data_scraped) == False].Ticker.tolist()
tickers_to_update = unanalyzed


# Download current data for all companies that are to be analyzed
current_data_for_all = download_current_data_for_all(unanalyzed)

data_available = current_data_for_all.Ticker.tolist()
y_sp_lst = [y_sp for ticker in data_available]
yrs_dg = [get_no_years(ccc, ticker) for ticker in data_available]
dgr1 = [get_dgr(ccc, ticker, 1) / 100 for ticker in data_available]
dgr3 = [get_dgr(ccc, ticker, 3) / 100 for ticker in data_available]
dgr5 = [get_dgr(ccc, ticker, 5) / 100 for ticker in data_available]
dgr10 = [get_dgr(ccc, ticker, 10) / 100 for ticker in data_available]

current_data_for_all[SP_YIELD] = y_sp_lst
current_data_for_all[YRS_DG] = yrs_dg
current_data_for_all[DGR1] = dgr1
current_data_for_all[DGR3] = dgr3
current_data_for_all[DGR5] = dgr5
current_data_for_all[DGR10] = dgr10

for ticker in tqdm.tqdm(tickers_to_update):
    print(f'--------------- Analyzing {ticker} ---------------')

    try:
        historical_data = load_historical_data(ticker)
    except FileNotFoundError:
        print(f"Data not downloaded for {ticker}, skipping")
        continue

    # report_month = historical_data[INCOME_STATEMENT].index[0].month
    # latest_year_reported = historical_data[INCOME_STATEMENT].index[0].year

    # current_year = datetime.date.today().year
    # current_month = datetime.date.today().month

    # if (latest_year_reported == current_year) or (current_year - latest_year_reported == 1 and current_month < report_month):
    #     # Report for current year has already been downloaded or has not been released yet
    #     pass
    # else:
    #     # We might have a newer report available, check
    #     latest_year_available = scraper.year_newest_report(ticker)
    #     if latest_year_available > latest_year_reported:
    #         # Newer report is available, re-download financials
    #         scraper.download_financial_data(ticker)

    # # Now we have the latest financial data downloaded

    # # Download the latest price chart
    # scraper.download_price_chart(ticker)

    # Download present data about the company

    # Perform the analysis
    current_data = current_data_for_all[current_data_for_all.Ticker == ticker]
    if current_data.empty:
        print(f"No current data available for {ticker}, skipping analysis")
        continue
    analysis = perform_analysis(historical_data, current_data)

    ist = historical_data[INCOME_STATEMENT]
    bsh = historical_data[BALANCE_SHEET]
    cfl = historical_data[CASH_FLOW]
    fra = historical_data[FINANCIAL_RATIOS]
    phi = historical_data[PRICE_HISTORY]
    dhi = historical_data[DIVIDEND_HISTORY]

    prices = analysis[PRICES]
    scores = analysis[SCORES]
    means = analysis[MEANS]

    # Update the entry in the list
    lst.loc[lst.Ticker == ticker,
            LAST_UPDATED] = datetime.date.today().strftime('%d-%b-%Y')
    lst.loc[lst.Ticker == ticker, ANNUAL_REPORT] = ist.index[0].strftime('%B')
    lst.loc[lst.Ticker == ticker, TOTAL_SCORE] = scores.sum(axis=1).iloc[0]
    lst.loc[lst.Ticker == ticker, 'Yrs Div Increase'] = current_data[YRS_DG]
    lst.loc[lst.Ticker == ticker, 'Avg Yield'] = prices[PYIELD][0]
    lst.loc[lst.Ticker == ticker, 'P/E Ratio'] = prices[PPE][0]
    lst.loc[lst.Ticker == ticker, 'P/EBIT'] = prices[PPEBIT][0]
    lst.loc[lst.Ticker == ticker, 'P/Op. CF'] = prices[PPOCF][0]
    lst.loc[lst.Ticker == ticker, 'P/FCF'] = prices[PPFCF][0]
    lst.loc[lst.Ticker == ticker, 'Gordon Growth'] = prices[PGGM][0]
    lst.loc[lst.Ticker == ticker, 'DCF'] = prices[PDCF][0]
    lst.loc[lst.Ticker == ticker, 'Avg Est. Price'] = prices[PEST][0]
    lst.loc[lst.Ticker == ticker, MIN_BUY] = analysis[MIN_BUY]
    lst.loc[lst.Ticker == ticker, MAX_BUY] = analysis[MAX_BUY]
    lst.loc[lst.Ticker == ticker,
            CURRENT_PRICE] = current_data[CURRENT_PRICE].iloc[0]
    lst.loc[lst.Ticker == ticker,
            CURRENT_YLD] = current_data[CURRENT_YLD].iloc[0]
    lst.loc[lst.Ticker == ticker, BETA] = current_data[BETA]
    lst.loc[lst.Ticker == ticker, DIVIDEND_SAFETY] = 99
    lst.loc[lst.Ticker == ticker, DGR3] = current_data[DGR3].iloc[0]
    lst.loc[lst.Ticker == ticker, LATEST_INCREASE] = analysis[LATEST_INCREASE]
    lst.loc[lst.Ticker == ticker,
            'Min Price Over/Under'] = current_data[CURRENT_PRICE].iloc[0] / analysis[MIN_BUY]
    lst.loc[lst.Ticker == ticker,
            'Max Price Over/Under'] = current_data[CURRENT_PRICE].iloc[0] / analysis[MAX_BUY]
    lst.loc[lst.Ticker == ticker, 'Sector'] = current_data[SECTOR]


print('--------------- END ---------------')


def write_lst_to_file(lst):
    writer = pd.ExcelWriter(file_path, engine='xlsxwriter',
                            datetime_format='d-mmm-yyyy',
                            date_format='d-mmm-yyyy')

    get_format_module().header_style = None

    lst.to_excel(
        writer, sheet_name="Analysis", startrow=6, startcol=1, index=False)

    sheet = writer.sheets["Analysis"]

    format_pct = writer.book.add_format({
        'num_format': '0.00%'
    })

    format_header = writer.book.add_format({
        'text_wrap': True,
        'bg_color': 'gray',
        'bold': True,
    })

    format_title = writer.book.add_format({
        'bold': True,
        'font_size': 36
    })

    format_two_dec = writer.book.add_format({
        'num_format': '0.00'
    })

    sheet.write('E4', 'List of companies', format_title)
    sheet.set_column('A:A', 3)
    sheet.set_column('B:B', 35.44)
    sheet.set_column('C:C', 6.11)
    sheet.set_column('D:D', 11.33)
    sheet.set_column('E:E', 9.00)
    sheet.set_column('F:F', 7.22)
    sheet.set_column('H:H', 7.89, format_two_dec)
    sheet.set_column('G:G', 7.67, format_two_dec)
    sheet.set_column('I:I', 8, format_two_dec)
    sheet.set_column('J:J', 7, format_two_dec)
    sheet.set_column('K:K', 7, format_two_dec)
    sheet.set_column('L:L', 6, format_two_dec)
    sheet.set_column('M:M', 6.56, format_two_dec)
    sheet.set_column('N:N', 6, format_two_dec)
    sheet.set_column('O:O', 6.89, format_two_dec)
    sheet.set_column('P:P', 11.78)
    sheet.set_column('Q:Q', 11.22)
    sheet.set_column('S:S', 6.56, format_pct)
    sheet.set_column('T:T', 4)
    sheet.set_column('U:U', 8.11)
    sheet.set_column('V:V', 7.22, format_pct)
    sheet.set_column('W:W', 10.33, format_pct)
    sheet.set_column('X:X', 10.33, format_pct)
    sheet.set_column('Y:Y', 7.56, format_pct)
    sheet.set_column('Z:Z', 6.22, format_pct)
    sheet.set_column('AA:AA', 6.22, format_pct)
    sheet.set_column('AB:AB', 20.33)

    sheet.set_row(6, None, format_header)

    writer.close()
