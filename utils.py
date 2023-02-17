
import datetime
import os
from copy import deepcopy

import numpy as np
import pandas as pd

from scoring import *
from constants import *
from scrape import Scraper


def get_ovy(years, dhi, phi, ist):
    highs = []
    for year in years:
        max_price = phi[phi.Date.dt.year == year].Close.max()
        try:
            dps = ist.loc[ist.index.year == year, 'Dividend Per Share'][0]
        except IndexError as e:
            print(
                f'Could not retrieve DPS for {year} from income statement. Attempting from dhi...')
            dps = dhi.groupby(dhi.Date.dt.year).sum().loc[year, 'Amount']
        max_yld = dps / max_price
        highs.append({
            'YEAR': year,
            'MAX_PRICE': max_price,
            'DIV_PAID': dps,
            'MAX_YLD': max_yld
        })
    return highs


def get_uvy(years, dhi, phi, ist):
    lows = []
    for year in years:
        min_price = phi[phi.Date.dt.year == year].Close.min()
        try:
            dps = ist.loc[ist.index.year == year, 'Dividend Per Share'][0]
        except IndexError as e:
            print(
                f'Could not retrieve DPS for {year} from income statement. Attempting from dhi...')
            dps = dhi.groupby(dhi.Date.dt.year).sum().loc[year, 'Amount']
        min_yld = dps / min_price
        lows.append({
            'YEAR': year,
            'MIN_PRICE': min_price,
            'DIV_PAID': dps,
            'MIN_YLD': min_yld})
    return lows


def validate_historical_data():
    for d in os.listdir(os.path.join(os.path.join(os.getcwd(), "data"))):
        financials_file = os.path.join(
            os.getcwd(), "data", d, f"{d.lower()}-financials.xlsx")
        dividends_file = os.path.join(
            os.getcwd(), "data", d, f"{d.lower()}-dividends.xlsx")
        chart_file = os.path.join(
            os.getcwd(), "data", d, f"{d.lower()}-chart.xlsx")

        financials_downloaded = os.path.exists(financials_file)
        dividends_downloaded = os.path.exists(dividends_file)
        chart_downloaded = os.path.exists(chart_file)
        all_good = financials_downloaded and dividends_downloaded and chart_downloaded

        if not all_good:
            print(f"-------------- {d} --------------")
            if not financials_downloaded:
                print("Financial data -- No ")
            if not dividends_downloaded:
                print("Dividend History -- No")
            if not chart_downloaded:
                print("Chart file -- No")

        if chart_downloaded:
            sz = os.path.getsize(chart_file)
            if sz / 1024 < 30:
                print(
                    f"Price chart of {d} incorrectly downloaded. Deleting...")
                os.remove(chart_file)


def load_historical_data(ticker):
    price_path = os.path.join(
        os.getcwd(), "data", ticker, f"{ticker.lower()}-chart.xlsx")

    price_history = pd.read_excel(price_path)
    price_history['Date'] = price_history['Date'].astype('datetime64')
    financials_path = os.path.join(
        os.getcwd(), "data", ticker, f"{ticker.lower()}-financials.xlsx")

    income_statement = pd.read_excel(
        financials_path, sheet_name='Income-Annual', index_col=0).T
    income_statement.index = income_statement.index.astype('datetime64[ns]')

    balance_sheet = pd.read_excel(
        financials_path, sheet_name='Balance-Sheet-Annual', index_col=0).T
    balance_sheet.index = balance_sheet.index.astype('datetime64[ns]')

    cash_flow = pd.read_excel(
        financials_path, sheet_name='Cash-Flow-Annual', index_col=0).T
    cash_flow.index = cash_flow.index.astype('datetime64[ns]')

    financial_ratios = pd.read_excel(
        financials_path, sheet_name='Ratios-Annual', index_col=0).T
    financial_ratios.index = financial_ratios.index.astype('datetime64[ns]')

    balance_sheet[DEBT_TO_CAPITAL] = balance_sheet[TOTAL_DEBT] / \
        (balance_sheet[TOTAL_ASSETS] -
         balance_sheet[TOTAL_LIABILITIES] + balance_sheet[TOTAL_DEBT])

    dividends_path = os.path.join(
        os.getcwd(), "data", ticker, f"{ticker.lower()}-dividends.xlsx")
    dividend_history = pd.read_excel(dividends_path)

    return {
        PRICE_HISTORY: price_history,
        DIVIDEND_HISTORY: dividend_history,
        INCOME_STATEMENT: income_statement,
        BALANCE_SHEET: balance_sheet,
        CASH_FLOW: cash_flow,
        FINANCIAL_RATIOS: financial_ratios
    }


def dgrN(dpy, year, n):
    try:
        b = dpy.loc[dpy.index == year, 'Amount'].iloc[0]
        a = dpy.loc[dpy.index == year - n, 'Amount'].iloc[0]
    except IndexError:
        return np.nan
    return (b/a) ** (1/n) - 1


def perform_analysis(historical_data, current_data):

    ist = historical_data[INCOME_STATEMENT]
    bsh = historical_data[BALANCE_SHEET]
    cfl = historical_data[CASH_FLOW]
    fra = historical_data[FINANCIAL_RATIOS]
    phi = historical_data[PRICE_HISTORY]
    dhi = historical_data[DIVIDEND_HISTORY]

    p_current = current_data[CURRENT_PRICE].iloc[0]
    y_current = current_data[CURRENT_YLD].iloc[0]
    y_sp_current = current_data[SP_YIELD].iloc[0]

    # Extract the last reported measures
    last_eps = ist[EPSB].iloc[0]
    last_ebit = ist[OPINC].iloc[0]
    last_shares = ist[SHARES].iloc[0]
    last_ocf = cfl[OCF].iloc[0]
    last_fcf = cfl[FCF].iloc[0]

    # Calculate historic means
    div_per_year = dhi.groupby(dhi.Date.dt.year).sum()
    price_per_year = phi.groupby(phi.Date.dt.year).mean()

    df = ist.groupby(ist.index.year).sum().join(
        price_per_year)[['Close', OPINC, SHARES, EPSB]]
    df = df.join(cfl.groupby(cfl.index.year).sum()[[FCF, OCF]])
    df = df.join(div_per_year)

    pe_hist = (df['Close'] / df[EPSB]).mean()
    pebit_hist = (df['Close'] / (df[OPINC] / df[SHARES])).mean()
    pocf_hist = (df['Close'] / (df[OCF] / df[SHARES])).mean()
    pfcf_hist = (df['Close'] / (df[FCF] / df[SHARES])).mean()
    y_hist = df['Yield'] = (df['Amount'] / df['Close']).mean()

    hist_means = pd.DataFrame.from_records([{
        YLD: y_hist,
        PE: pe_hist,
        PEBIT: pebit_hist,
        POCF: pocf_hist,
        PFCF: pfcf_hist,
    }])

    # Price using yield
    p_yield = p_current / max(0.8, y_hist / y_current)

    # Price using P/E Ratio
    denominator = (p_current / last_eps) / pe_hist
    if denominator < 0:
        denominator = 1.5
    p_pe = p_current / denominator

    # Price using P/EBIT Ratio
    current_pebit = p_current / (last_ebit / last_shares)
    denominator = current_pebit / pebit_hist
    if denominator < 0:
        denominator = 1.5
    p_ebit = p_current / denominator

    # Price using P/OpCF Ratio
    current_pocf = p_current / (last_ocf / last_shares)
    denominator = current_pocf / pocf_hist
    if denominator < 0:
        denominator = 1.5
    p_ocf = p_current / denominator

    # Price using P/FCF Ratio
    current_pfcf = p_current / (last_fcf / last_shares)
    denominator = current_pfcf / pfcf_hist
    if denominator < 0:
        denominator = 1.5
    p_fcf = p_current / denominator

    # Price using Gordon Growth Model
    p_ggm = 0

    # Price using Discounted Cash Flow
    p_dcf = 0

    # Average Estimated Price
    p_est = (p_yield + p_pe + p_ebit + p_ocf + p_fcf + p_ggm + p_dcf) / 7

    prices = pd.DataFrame.from_records([{
        PYIELD: p_yield,
        PPE: p_pe,
        PPEBIT: p_ebit,
        PPOCF: p_ocf,
        PPFCF: p_fcf,
        PGGM: p_ggm,
        PDCF: p_dcf,
        PEST: p_est
    }])

    # Scores
    yrs_dg = current_data[YRS_DG].iloc[0]
    yrs_dg_score = get_yrs_dg_score(yrs_dg)

    yld_score = get_yield_score(y_current, y_sp_current)

    # TODO - Take DGR1, DGR3, DGR5, DGR10 into account
    dgr10 = current_data[DGR10].iloc[0]
    dgr5 = current_data[DGR5].iloc[0]
    dgr3 = current_data[DGR3].iloc[0]
    dgr1 = current_data[DGR1].iloc[0]
    dgr_score = get_dgr_score(dgr10, y_current)

    # Compute EPR and FCFPR
    eps_div_fcf = ist[[EPSB, DPS]].join(cfl[FCFPS])
    eps_div_fcf[FCFPR] = eps_div_fcf[DPS] / eps_div_fcf[FCFPS]
    eps_div_fcf[EPR] = eps_div_fcf[DPS] / eps_div_fcf[EPSB]

    # TODO - Take REIT/not into account
    eps_pr_score = get_epr_score(eps_div_fcf[[EPSB, EPR]])
    fcf_pr_score = get_fcfpr_score(eps_div_fcf[[FCFPS, FCFPR]])

    op_margin = ist[OPM].iloc[0]
    op_margin_score = get_op_margin_score(op_margin)

    roe = fra[ROE].iloc[0]
    roe_score = get_roe_score(roe)

    dtc = bsh[DEBT_TO_CAPITAL].iloc[0]
    dtc_score = get_dtc_score(dtc)

    # TODO - How ?
    shares = ist[SHARES]
    shares_score = get_shares_score(shares)

    # TODO - Where to get Credit Rating?
    credit_rating = current_data[CREDIT_RATING].iloc[0]
    credit_rating_score = get_credit_rating_score(credit_rating)

    scores = pd.DataFrame.from_records([{
        YRS_DG: yrs_dg_score,
        YLD: yld_score,
        DGR: dgr_score,
        EPR: eps_pr_score,
        FCFPR: fcf_pr_score,
        DEBT_TO_CAPITAL: dtc_score,
        ROE: roe_score,
        OPM: op_margin_score,
        SHARES: shares_score,
        CREDIT_RATING: credit_rating_score
    }])

    return {
        PRICES: prices,
        SCORES: scores,
        MEANS: hist_means,
        MAX_BUY: 1,
        MIN_BUY: 1,
        LATEST_INCREASE: (lambda x: x[x.gt(0)])(
            dhi['Amount'].iloc[::-1].pct_change()).iloc[-1],
        YRS_DG: yrs_dg
    }


def write_historical_data(writer, historical_data):
    ist = historical_data[INCOME_STATEMENT].T
    bsh = historical_data[BALANCE_SHEET].T
    cfl = historical_data[CASH_FLOW].T
    fra = historical_data[FINANCIAL_RATIOS].T

    ist_yoy = ist.iloc[::-1].pct_change()
    bsh_yoy = bsh.iloc[::-1].pct_change()
    cfl_yoy = cfl.iloc[::-1].pct_change()

    # Write financial statements
    row = 30
    ist.to_excel(
        writer, sheet_name="Analysis", startrow=row, startcol=1, index=True)
    row += ist.shape[0] + 2 + 1

    sheet = writer.sheets["Analysis"]
    sheet.write(30, 0, 'Income Statement')
    sheet.write(row, 0, 'Income Statement (YoY)')
    ist_yoy.to_excel(
        writer, sheet_name="Analysis", startrow=row, startcol=1, index=True)
    row += ist_yoy.shape[0] + 4 + 1

    sheet.write(row, 0, 'Balance Sheet')
    bsh.to_excel(writer, sheet_name="Analysis",
                 startrow=row, startcol=1, index=True)
    row += bsh.shape[0] + 2 + 1

    sheet.write(row, 0, 'Balance Sheet (YoY)')
    bsh_yoy.to_excel(
        writer, sheet_name="Analysis", startrow=row, startcol=1, index=True)
    row += bsh_yoy.shape[0] + 4 + 1

    sheet.write(row, 0, 'Cash Flow Statement')
    cfl.to_excel(writer, sheet_name="Analysis",
                 startrow=row, startcol=1, index=True)
    row += cfl.shape[0] + 2 + 1

    sheet.write(row, 0, 'Cash Flow Statement (YoY)')
    cfl_yoy.to_excel(writer, sheet_name="Analysis",
                     startrow=row, startcol=1, index=True)
    row += cfl_yoy.shape[0] + 3 + 1

    sheet.write(row, 0, 'Financial Ratios & Other')
    fra.to_excel(
        writer, sheet_name="Analysis", startrow=row, startcol=1, index=True)

    sheet.set_column('A:A', 25)
    sheet.set_column('B:B', 25)


def createDycCharts(writer):
    return


def write_prices(writer, current_data, analysis):

    prices = analysis[PRICES]
    means = analysis[MEANS]

    format_two_dec = writer.book.add_format({
        'num_format': '0.00',
        'border': 1
    })

    format_pct = writer.book.add_format({
        'num_format': '0.00%',
        'border': 1
    })

    format_b_gray = writer.book.add_format({
        'bg_color': 'gray',
        'bold': True,
        'border': 1

    })
    format_td_b_gray = writer.book.add_format({
        'num_format': '0.00',
        'bg_color': 'gray',
        'bold': True,
        'border': 1
    })
    format_border = writer.book.add_format({
        'border': 1
    })

    sheet = writer.sheets['Analysis']

    c_name = current_data[NAME].iloc[0]
    ticker = current_data[TICKER].iloc[0]
    formatted_dt = datetime.date.today().strftime('%d.%m.%Y')
    rate_card_title = f"{c_name}({ticker}) - RATE CARD (last updated {formatted_dt})"
    sheet.merge_range('K4:N4', rate_card_title, format_b_gray)

    sheet.write('K5', 'Current Div/Share', format_border)
    sheet.write('L5', current_data[CURRENT_YLD].iloc[0], format_two_dec)

    sheet.write('K6', 'S&P 500 Yield', format_border)
    sheet.write('L6', current_data[SP_YIELD].iloc[0], format_pct)

    sheet.write('K7', 'Beta', format_border)
    sheet.write('L7', current_data[BETA].iloc[0], format_two_dec)

    sheet.write('K8', 'Avg Yield', format_border)
    sheet.write('L8', means[YLD], format_pct)

    sheet.write('K9', 'Price using Yield', format_border)
    sheet.write('L9', prices[PYIELD], format_two_dec)

    sheet.write('K10', 'Avg P/E Ratio', format_border)
    sheet.write('L10', means[PE], format_two_dec)

    sheet.write('K11', 'Price using P/E', format_border)
    sheet.write('L11', prices[PPE], format_two_dec)

    sheet.write('K12', 'Avg P/EBIT', format_border)
    sheet.write('L12', means[PEBIT], format_two_dec)

    sheet.write('K13', 'Price using P/EBIT', format_border)
    sheet.write('L13', prices[PPEBIT], format_two_dec)

    sheet.write('K14', 'Avg P/OpCF', format_border)
    sheet.write('L14', means[POCF], format_two_dec)

    sheet.write('K15', 'Price using P/OpCF', format_border)
    sheet.write('L15', prices[PPOCF], format_two_dec)

    sheet.write('K16', 'Avg P/FCF', format_border)
    sheet.write('L16', means[PFCF], format_two_dec)

    sheet.write('K17', 'Price using P/FCF', format_border)
    sheet.write('L17', prices[PPFCF], format_two_dec)

    sheet.write('K18', 'Gordon Growth Model', format_border)
    sheet.write('L18', prices[PGGM], format_two_dec)

    sheet.write('K19', 'Discounted Cash Flow', format_border)
    sheet.write('L19', prices[PDCF], format_two_dec)

    sheet.write('K20', 'Avg. Estimated Price', format_b_gray)
    sheet.write('L20', prices[PEST], format_td_b_gray)

    sheet.set_column('K:K', 18.78)


def write_scores(writer, historical_data, current_data, analysis):
    sheet = writer.sheets["Analysis"]

    ist = historical_data[INCOME_STATEMENT]
    bsh = historical_data[BALANCE_SHEET]
    fra = historical_data[FINANCIAL_RATIOS]

    scores = analysis[SCORES]

    format_pct = writer.book.add_format({
        'num_format': '0.00%',
        'border': 1
    })
    format_b_gray = writer.book.add_format({
        'bg_color': 'gray',
        'bold': True,
        'border': 1
    })
    format_border = writer.book.add_format({
        'border': 1
    })

    sheet.write('O4', 'Score', format_b_gray)
    sheet.write('M5', 'Yrs of Div Growth', format_border)

    sheet.write('N5', analysis[YRS_DG], format_border)
    sheet.write('O5', scores[YRS_DG], format_border)

    sheet.write('M6', 'Current Yield')
    sheet.write('N6', current_data[CURRENT_YLD].iloc[0], format_pct)
    sheet.write('O6', scores[YLD], format_border)

    sheet.write('M7', 'DGR10', format_border)
    sheet.write('N7', current_data[DGR10].iloc[0], format_pct)

    sheet.write('M8', 'DGR5', format_border)
    sheet.write('N8', current_data[DGR5].iloc[0], format_pct)

    sheet.write('M9', 'DGR3', format_border)
    sheet.write('N9', current_data[DGR3].iloc[0], format_pct)

    sheet.write('M10', 'DGR1', format_border)
    sheet.write('N10', current_data[DGR1].iloc[0], format_pct)

    sheet.write('M11', 'Last div Increase', format_border)
    sheet.write('N11', analysis[LATEST_INCREASE], format_pct)

    sheet.merge_range('O7:O11', scores[DGR], format_border)

    # TODO - ???
    sth = 'sth'
    sheet.write('M12', 'EPS & Payout R', format_border)
    sheet.write('N12', sth, format_border)
    sheet.write('O12', scores[EPR], format_border)

    sheet.write('M13', 'FCF & Payout R', format_border)
    sheet.write('N13', sth, format_border)
    sheet.write('O13', scores[FCFPR], format_border)

    sheet.write('M14', 'Debt/Capital', format_border)
    sheet.write('N14', bsh[DEBT_TO_CAPITAL].iloc[0], format_pct)
    sheet.write('O14', scores[DEBT_TO_CAPITAL], format_border)

    sheet.write('M15', 'ROE', format_border)
    sheet.write('N15', fra[ROE].iloc[0], format_pct)
    sheet.write('O15', scores[ROE], format_border)

    sheet.write('M16', 'Op. Margin', format_border)
    sheet.write('N16', ist[OPM].iloc[0], format_pct)
    sheet.write('O16', scores[OPM], format_border)

    sheet.write('M17', '# Shares', format_border)
    sheet.write('N17', sth, format_border)
    sheet.write('O17', scores[SHARES], format_border)

    sheet.write('M18', 'Credit Rating', format_border)
    sheet.write('N18', current_data[CREDIT_RATING].iloc[0])
    sheet.write('O18', scores[CREDIT_RATING], format_border)

    sheet.merge_range('M19:N20', 'TOTAL Score (Max 30)', format_b_gray)
    sheet.merge_range('O19:O20', scores.sum(axis=1).iloc[0], format_border)

    sheet.set_column('M:M', 17.33)


def create_charts(writer):
    # Profits Chart
    sheet = writer.sheets['Analysis']

    sheet.merge_range('A119:M121', 'GRAPHS AREA')
    chart = writer.book.add_chart({'type': 'line'})
    chart.add_series({'values': '=Analysis!$C$34:$L$34',
                      'categories': 'Analysis!$C$31:$L$31',
                      'name': 'Gross Profit'
                      })
    chart.add_series({'values': '=Analysis!$C$37:$L$37',
                      'categories': 'Analysis!$C$31:$L$31',
                      'name': 'Operating Income (EBIT)'})
    chart.add_series({'values': '=Analysis!$C$40:$L$40',
                      'categories': 'Analysis!$C$31:$L$31',
                      'name': 'Net Income'})

    chart.set_title({
        'name': 'Profits',
    })
    sheet.insert_chart('A123', chart)

    # Profitability margins
    chart = writer.book.add_chart({'type': 'line'})
    chart.add_series({'values': '=Analysis!$C$35:$L$35',
                      'categories': 'Analysis!$C$31:$L$31',
                      'name': 'Gross Margin'
                      })
    chart.add_series({'values': '=Analysis!$C$38:$L$38',
                      'categories': 'Analysis!$C$31:$L$31',
                      'name': 'Operating Margin'})
    chart.add_series({'values': '=Analysis!$C$41:$L$41',
                      'categories': 'Analysis!$C$31:$L$31',
                      'name': 'Net Margin'})

    chart.set_title({
        'name': 'Profitability margins',
    })
    sheet.insert_chart('F123', chart)

    # Revenue vs COGS
    chart = writer.book.add_chart({'type': 'line'})
    chart.add_series({'values': '=Analysis!$C$32:$L$32',
                      'categories': 'Analysis!$C$31:$L$31',
                      'name': 'Revenue'
                      })
    chart.add_series({'values': '=Analysis!$C$33:$L$33',
                      'categories': 'Analysis!$C$31:$L$31',
                      'name': 'Total COGS'})

    chart.set_title({
        'name': 'Revenue vs COGS',
    })
    sheet.insert_chart('A138', chart)

    # Cash Flows
    chart = writer.book.add_chart({'type': 'line'})
    chart.add_series({'values': '=Analysis!$C$85:$L$85',
                      'categories': 'Analysis!$C$31:$L$31',
                      'name': 'Operating Cash Flow'
                      })
    chart.add_series({'values': '=Analysis!$C$86:$L$86',
                      'categories': 'Analysis!$C$31:$L$31',
                      'name': 'Financing Cash Flow'})
    chart.add_series({'values': '=Analysis!$C$88:$L$88',
                      'categories': 'Analysis!$C$31:$L$31',
                      'name': 'Investing  Cash Flow'})

    chart.set_title({
        'name': 'Cash Flows',
    })
    sheet.insert_chart('F138', chart)

    # E/FCF/Div per Share
    chart = writer.book.add_chart({'type': 'line'})
    chart.add_series({'values': '=Analysis!$C$42:$L$42',
                      'categories': 'Analysis!$C$31:$L$31',
                      'name': 'EPS'
                      })
    chart.add_series({'values': '=Analysis!$C$90:$L$90',
                      'categories': 'Analysis!$C$31:$L$31',
                      'name': 'FCF Per Share'})
    chart.add_series({'values': '=Analysis!$C$92:$L$92',
                      'categories': 'Analysis!$C$31:$L$31',
                      'name': 'Dividend Per Share'})

    chart.set_title({
        'name': 'E/FCF/Div Per Share',
    })
    sheet.insert_chart('A153', chart)

    # EPS/FCF Payout Ratio
    chart = writer.book.add_chart({'type': 'line'})
    chart.add_series({'values': '=Analysis!$C$108:$L$108',
                      'categories': 'Analysis!$C$31:$L$31',
                      'name': 'EPS Payout Ratio'
                      })
    chart.add_series({'values': '=Analysis!$C$109:$L$109',
                      'categories': 'Analysis!$C$31:$L$31',
                      'name': 'FCF Payout Ratio'})

    chart.set_title({
        'name': 'EPS/FCF Payout Ratio',
    })
    sheet.insert_chart('F153', chart)

    # Shares outstanding
    chart = writer.book.add_chart({'type': 'line'})
    chart.add_series({'values': '=Analysis!$C$43:$L$43',
                      'categories': 'Analysis!$C$31:$L$31',
                      'name': 'Shares outstanding'
                      })

    chart.set_title({
        'name': 'Shares outstanding',
    })
    sheet.insert_chart('A168', chart)

    # Debt & Return
    chart = writer.book.add_chart({'type': 'line'})
    chart.add_series({'values': '=Analysis!$C$69:$L$69',
                      'categories': 'Analysis!$C$31:$L$31',
                      'name': 'Debt/Capital'
                      })
    chart.add_series({'values': '=Analysis!$C$111:$L$111',
                      'categories': 'Analysis!$C$31:$L$31',
                      'name': 'Return On Equity (ROE)'})

    chart.set_title({
        'name': 'Debt & Return',
    })
    sheet.insert_chart('F168', chart)

    # Ratios
    chart = writer.book.add_chart({'type': 'line'})
    chart.add_series({'values': '=Analysis!$C$110:$L$110',
                      'categories': 'Analysis!$C$31:$L$31',
                      'name': 'P/E Ratio'
                      })
    chart.add_series({'values': '=Analysis!$C$112:$L$112',
                      'categories': 'Analysis!$C$31:$L$31',
                      'name': 'P/EBIT Ratio'})
    chart.add_series({'values': '=Analysis!$C$113:$L$113',
                      'categories': 'Analysis!$C$31:$L$31',
                      'name': 'P/OpCF Ratio'})
    chart.add_series({'values': '=Analysis!$C$114:$L$114',
                      'categories': 'Analysis!$C$31:$L$31',
                      'name': 'P/FCF Ratio'})

    chart.set_title({
        'name': 'Ratios',
    })

    sheet.insert_chart('A183', chart)


def create_company_list(file_path):
    writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
    sheet = writer.sheets['Analysis']

    sheet.merge_range('B8', 'Name')
    sheet.merge_range('C8', 'Ticker')
    sheet.merge_range('D8', 'Last Updated')
    sheet.merge_range('E8', 'Annual Report')
    sheet.merge_range('F8', 'Yrs Div Increase')
    sheet.merge_range('G8', 'Score (Max 30)')
    sheet.merge_range('H7:N7', 'Estimated Share Price')
    sheet.merge_range('H8', 'Avg Yield')
    sheet.merge_range('I8', 'P/E Ratio')
    sheet.merge_range('J8', 'P/EBIT')
    sheet.merge_range('K8', 'P/Op. Cf')
    sheet.merge_range('L8', 'P/FCF')
    sheet.merge_range('M8', 'Gordon Growth')
    sheet.merge_range('N8', 'DCF')
    sheet.merge_range('O8', 'Avg. Est Price')
    sheet.merge_range('P8', 'Min Buy Price')
    sheet.merge_range('Q8', 'Max Buy Price')
    sheet.merge_range('R8', 'Current Price')
    sheet.merge_range('S8', 'Current Yield')
    sheet.merge_range('T8', 'Beta')
    sheet.merge_range('U8', 'Dividend Safety')
    sheet.merge_range('V8', 'DGR3')
    sheet.merge_range('W8', 'Latest Increase')
    sheet.merge_range('X8', 'Min Price Over/Under')
    sheet.merge_range('Y8', 'Max Price Over/Under')

    writer.close()


def get_list_entry(historical_data, current_data, analysis):
    # Company does not exist, add it to the df
    scores = analysis[SCORES]
    prices = analysis[PRICES]
    ist = historical_data[INCOME_STATEMENT]
    dhi = historical_data[DIVIDEND_HISTORY]

    dpy = dhi.groupby(dhi.Date.dt.year).sum()

    # TODO - ???
    current_year = 2022

    # TODO pd.Series
    new_entries = {
        NAME: current_data[NAME],
        TICKER: current_data[TICKER],
        LAST_UPDATED: datetime.date.today().strftime('%d-%b-%Y'),
        ANNUAL_REPORT: ist.index[0].strftime('%B'),
        'Yrs Div Increase': analysis[LATEST_INCREASE],
        TOTAL_SCORE: scores.sum(axis=1).iloc[0],
        'Avg Yield': prices[PYIELD],
        'P/E Ratio': prices[PPE],
        'P/EBIT': prices[PPEBIT],
        'P/Op. Cf': prices[PPOCF],
        'P/FCF': prices[PPFCF],
        'Gordon Growth': prices[PGGM],
        'DCF': prices[PDCF],
        'Avg Est. Price': prices[PEST],
        'Min Buy Price': 'min_buy',
        'Max Buy Price': 'max_buy',
        'Current Price': current_data[CURRENT_PRICE],
        'Current Yield': current_data[CURRENT_YLD],
        'Beta': current_data[BETA],
        'Dividend Safety': 'dsafety',
        'DGR3': dgrN(dpy, current_year, 3),
        'Latest Increase': analysis[LATEST_INCREASE],
        'Min Price Over/Under': 'minpct',
        'Max Price Over/Under': 'maxpct',
    }

    return new_entries


def update_company_list_entry(ticker, new_entries):

    file_path = os.path.join(os.getcwd(), "company-list.xlsx")

    if not os.path.exists(file_path):
        create_company_list(file_path)
    # Read the entire company list and search for the company
    df = pd.read_excel(file_path, header=7).iloc[:, 1:]
    if ticker in df['Ticker'].values:
        df[df['Ticker'] == ticker] = new_entries
    else:
        df.append(new_entries, ignore_index=True)

    # Write the DF back to the file
    writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
    df.to_excel(writer, sheet_name="Analysis",
                startrow=6, startcol=1, index=True)
    writer.close()


def is_data_scraped(ticker):
    folder = os.path.join(os.getcwd(), "data", ticker)
    if not os.path.exists(os.path.join(folder, f"{ticker.lower()}-chart.xlsx")):
        return False
    if not os.path.exists(os.path.join(folder, f"{ticker.lower()}-financials.xlsx")):
        return False
    if not os.path.exists(os.path.join(folder, f"{ticker.lower()}-dividends.xlsx")):
        return False
    return True


def update_historical_data(ticker):
    return


def update_company(ticker, y_sp):
    if is_data_scraped(ticker):
        if newerReportAvailable(ticker):
            update_historical_data(ticker)
    else:
        # print('Scraping company data ... ', end='')
        scraper.download_historical_data(ticker)

    # Load the latest current data, and the historical data
    company_info = scraper.download_company_info(ticker)
    current_data = deepcopy(company_info)
    current_data[SP_YIELD] = y_sp
    historical_data = load_historical_data(ticker)

    # Now we have the newest data downloaded (part in RAM, part in files)

    # Perform analysis
    analysis = perform_analysis(historical_data, current_data)

    # Write newest analysis into an Excel file for analysis)
    folder = os.path.join(os.getcwd(), "analysis")
    if not os.path.exists(folder):
        os.makedirs(folder)
    file_path = os.path.join(folder, f"{ticker.lower()}.xlsx")
    writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
    write_historical_data(writer, historical_data)
    write_prices(writer, current_data, analysis)
    write_scores(writer, historical_data, current_data, analysis)
    create_charts(writer)
    createDycCharts(writer)
    writer.close()

    # Get the entry values - formatted
    entry = get_list_entry(historical_data, current_data, analysis)

    # Update the entry in the database
    update_company_list_entry(entry)


def read_tickers(ticker_list_file):
    with open(os.path.join(os.getcwd(), ticker_list_file)) as f:
        tickers = f.readlines()
    return tickers[0].split(' ')


def get_unscraped_tickers():
    company_list = pd.read_excel(os.path.join(
        os.getcwd(), "company-list.xlsx"), header=7).iloc[:, 1:]
    unscraped_tickers = company_list[company_list.Ticker.apply(
        is_data_scraped) == False].Ticker.tolist()
    return unscraped_tickers


def get_champions():
    champions = pd.read_excel(
        "CCC List.xlsx", sheet_name="Champions", header=2)
    return champions


def get_contenders():
    contenders = pd.read_excel(
        "CCC List.xlsx", sheet_name="Contenders", header=2)
    return contenders


def get_challengers():
    challengers = pd.read_excel(
        "CCC List.xlsx", sheet_name="Challengers", header=2)
    return challengers


def get_no_years(ccc, ticker):
    try:
        no_years = ccc.loc[ccc.Symbol == ticker, 'No Years'].iloc[0]
    except IndexError:
        print(f"{ticker} could not be found in CCC List. Assuming 0 years")
        no_years = 0
    return no_years


def get_dgr(ccc, ticker, n):
    try:
        dgr = ccc.loc[ccc.Symbol == ticker, f"DGR {n}Y"].iloc[0]
    except IndexError:
        print(f"{ticker} could not be found in CCC List. No DGR available")
        dgr = 0
    return dgr


def get_ccc():
    champions = get_champions()
    contenders = get_contenders()
    challengers = get_challengers()
    ccc = champions.append(contenders).append(challengers)
    return ccc
