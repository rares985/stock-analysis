import matplotlib.pyplot as plt

from constants import *


def get_op_margin_score(op_margin):
    if op_margin > OPM_HI:
        return 3
    elif op_margin > OPM_MID:
        return 2
    elif op_margin > OPM_LO:
        return 1
    return -1


def get_roe_score(roe):
    if roe > ROE_HI:
        return 3
    elif roe > ROE_MID:
        return 2
    elif roe > ROE_LO:
        return 1
    return -1


def get_dtc_score(dtc):
    if dtc < DTC_HI:
        return 3
    elif dtc < DTC_MID:
        return 2
    elif dtc < DTC_LO:
        return 1
    return -1


def get_yrs_dg_score(yrs_dg):
    if yrs_dg > YRS_DG_HI:
        return 3
    elif yrs_dg > YRS_DG_MID:
        return 2
    elif yrs_dg > YRS_DG_LO:
        return 1
    return -1


def get_yield_score(yld, sp):
    if yld > YLD_HI * sp:
        return 3
    elif yld > YLD_LO * sp:
        return 2
    else:
        return 1


def get_dgr_score(dgr, yld):
    if yld > 4:
        if dgr > 7:
            return 3
        elif dgr > 4:
            return 2
        elif dgr > 1:
            return 1
        else:
            return -1
    elif yld > 2:
        if dgr > 10:
            return 3
        elif dgr > 5:
            return 2
        elif dgr > 2:
            return 1
        else:
            return -1
    else:
        if dgr > 13:
            return 3
        elif dgr > 9:
            return 2
        elif dgr > 5:
            return 1
        else:
            return -1


def get_epr_score():
    # TODO - how
    return 0


def get_fcfpr_score():
    # TODO - how
    return 0


def get_shares_score(shares):
    # TODO - how
    return 0


def get_credit_rating_score(rating):
    if rating in ['AAA, AA+', 'AA', 'AA']:
        return 3
    elif rating in ['A+', 'A', 'A-', 'BBB+']:
        return 2
    elif rating in ['BBB', 'BBB-']:
        return 1
    else:
        return -1


# all input dfs must be with year as index, increasing, and transposed


def plot_advanced_data(df, ticker):
    df = df.iloc[::-1]
    df = df.set_index('Dates')

    fig, ax = plt.subplots(ncols=2, nrows=5, figsize=(30, 20))
    fig.suptitle(ticker, size=50)
    # Profits
    ax[0, 0].plot(df['Gross Profit'], lw=2,
                  marker='.', markersize=10, label="Gross Profit")
    ax[0, 0].plot(df['Operating Income'], lw=2,
                  marker='.', markersize=10, label="Operating Income")
    ax[0, 0].plot(df['Net Income'], lw=2,
                  marker='.', markersize=10, label="Net Income")
    ax[0, 0].legend()
    ax[0, 0].set_title('Profits')

    # Profitability margins
    ax[0, 1].plot(df['Gross Margin'], lw=2,
                  marker='.', markersize=10, label="Gross Margin")
    ax[0, 1].plot(df['Operating Margin'], lw=2,
                  marker='.', markersize=10, label="Operating Income")
    ax[0, 1].plot(df['Net Profit Margin'], lw=2,
                  marker='.', markersize=10, label="Net Profit Margin")
    ax[0, 1].legend()
    ax[0, 1].set_title('Profitability margins')

    # Revenue vs COGS
    ax[1, 0].plot(df['Revenue'], lw=2,
                  marker='.', markersize=10, label="Revenue")
    ax[1, 0].plot(df['Cost Of Goods Sold'], lw=2,
                  marker='.', markersize=10, label="Cost Of Goods Sold")
    ax[1, 0].legend()
    ax[1, 0].set_title('Revenue vs Total COGS')

    # Cash Flows
    ax[1, 1].plot(df['Cash Flow From Operating Activities'], lw=2,
                  marker='.', markersize=10, label="Operating Cash Flow")
    ax[1, 1].plot(df['Cash Flow From Investing Activities'], lw=2,
                  marker='.', markersize=10, label="Investing Cash Flow")
    ax[1, 1].plot(df['Cash Flow From Financial Activities'], lw=2,
                  marker='.', markersize=10, label="Financing Cash Flow")
    ax[1, 1].legend()
    ax[1, 1].set_title('Cash Flows')

    ax[2, 0].plot(df['Basic EPS'], lw=2,
                  marker='.', markersize=10, label="EPS (basic) ")
    # TODO
    # ax[2, 0].plot(df['Dividends Per Share'], lw=2,
    #                        marker='.', markersize=10, label="Dividends Per Share")
    ax[2, 0].plot(df['Free Cash Flow Per Share'], lw=2,
                  marker='.', markersize=10, label="Free Cash Flow Per Share")
    ax[2, 0].legend()
    ax[2, 0].set_title('E/FCF/Div per Share')

    # EPS / FCF Payout Ratio - TODO
    # ax[2, 1].plot(df['Earnings Payout Ratio'], lw=2,
    #                         marker='.', markersize=10, label="Earnings Payout Ratio")
    # ax[2, 1].plot(df['Free Cash Flow Per Share'], lw=2,
    #                        marker='.', markersize=10, label="FCF Payout Ratio")
    ax[2, 1].legend()
    ax[2, 1].set_title('E/FCF Payout Ratio')

    # Shares outstanding
    ax[3, 0].plot(df['Basic Shares Outstanding'], lw=2,
                  marker='.', markersize=10, label="# Shares Outstanding (basic) ")
    ax[3, 0].legend()
    ax[3, 0].set_title('# Shares outstanding')

    ax[3, 1].plot(df['Long-term Debt / Capital'], lw=2,
                  marker='.', markersize=10, label="Debt/Capital")
    ax[3, 1].plot(df['ROE - Return On Equity'], lw=2,
                  marker='.', markersize=10, label="Return On Equity (ROE)")
    ax[3, 1].legend()
    ax[3, 1].set_title('Debt/Capital and ROE')

    # Price ratios
    # ax[4, 0].plot(df['P/E Ratio'], lw=2,
    #                        marker='.', markersize=10, label="P/E Ratio")
    # ax[4, 0].plot(df['P/EBIT Ratio'], lw=2,
    #                        marker='.', markersize=10, label="P/EBIT Ratio")
    # ax[4, 0].plot(df['P/OpCF Ratio'], lw=2,
    #                        marker='.', markersize=10, label="P/OpCF Ratio")
    # ax[4, 0].plot(df['P/FCF Ratio'], lw=2,
    #                        marker='.', markersize=10, label="P/FCF Ratio")
    ax[4, 0].legend()
    ax[4, 0].set_title('Price ratios')

    # DYC
    # ax[4, 1].plot(df['Price'], lw=2,
    #                        marker='.', markersize=10, label="Price")
    # ax[4, 1].plot(df['Overvalued Yield'], lw=2,
    #                        marker='.', markersize=10, label="Overvalued Yield")
    # ax[4, 1].plot(df['Undervalued Yield'], lw=2,
    #                        marker='.', markersize=10, label="Undervalued Yield")
    # ax[4, 1].plot(df['75th Percentile'], lw=2,
    #                        marker='.', markersize=10, label="75th Percentile (25% to Overvalued)")
    # ax[4, 1].plot(df['25th Percentile'], lw=2,
    #                        marker='.', markersize=10, label="25th Percentile (25% to Undervalued)")
    ax[4, 1].legend()
    ax[4, 1].set_title('Dividend Yield Channels')

    return fig


def plot_data(complete, ticker):
    # managing plots
    fig1, f1_axes = plt.subplots(ncols=2, nrows=2, figsize=(30, 20))
    fig1.suptitle(ticker, size=50)
    f1_axes[0, 0].plot(complete['Revenue'], lw=2,
                       marker='.', markersize=10, label="Revenue")
    f1_axes[0, 0].plot(complete['Gross Profit'], lw=2,
                       marker='.', markersize=10, label="Gross Profit")
    f1_axes[0, 0].plot(complete['Net Income'], lw=2,
                       marker='.', markersize=10, label="Net Income")
    f1_axes[0, 0].plot(complete['EBITDA'], lw=2,
                       marker='.', markersize=10, label="EBITDA")
    f1_axes[0, 0].plot(complete['Total Assets'], lw=2,
                       marker='.', markersize=10, label="Total Assets")
    f1_axes[0, 0].plot(complete['Total Liabilities'], lw=2,
                       marker='.', markersize=10, label="Total Liabilities")
    f1_axes[0, 0].plot(complete['Total Depreciation And Amortization - Cash Flow'],
                       lw=2, marker='.', markersize=10, label="Cash Flow")
    f1_axes[0, 0].plot(complete['Net Cash Flow'], lw=2,
                       marker='.', markersize=10, label="Net Cash Flow")
    f1_axes[0, 1].plot(complete['EPS - Earnings Per Share'],
                       lw=2, marker='.', markersize=10, label="EPS")
    f1_axes[1, 0].plot(complete['ROE - Return On Equity'],
                       lw=2, marker='.', markersize=10, label="ROE")
    f1_axes[1, 0].plot(complete['ROA - Return On Assets'],
                       lw=2, marker='.', markersize=10, label="ROA")
    f1_axes[1, 0].plot(complete['ROI - Return On Investment'],
                       lw=2, marker='.', markersize=10, label="ROI")
    f1_axes[1, 1].plot(complete['Shares Outstanding'], lw=2,
                       marker='.', markersize=10, label="Shares Outstanding")
    f1_axes[0, 0].legend()
    f1_axes[0, 0].invert_xaxis()
    f1_axes[0, 1].legend()
    f1_axes[0, 1].invert_xaxis()
    f1_axes[1, 0].legend()
    f1_axes[1, 0].invert_xaxis()
    f1_axes[1, 1].legend()
    f1_axes[1, 1].invert_xaxis()

    return fig1
