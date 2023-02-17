# from scoring import *
from utils import load_historical_data


def plotYieldChannels(dphi, ovy, uvy):
    fig, ax = plt.subplots()
    ax.plot(dphi.Date, dphi.Close)
    twin = ax.twinx()
    twin2 = ax.twinx()

    ylds = [ovy, uvy]
    color = 'tab:red'
    for idx, yld in enumerate(ylds):
        twin.plot(dphi.Date, dphi.Amount / yld, color=color)
    twin.plot(dphi.Date, dphi['25% to OVV'],
              color='tab:green', linestyle='dashed')
    twin.plot(dphi.Date, dphi['25% to UVV'],
              color='tab:green', linestyle='dashed')
    # twin2.plot(dphi.Date, dphi.Yield, color='tab:green')
    plt.show()


ticker = 'APOG'
high_years = [2007, 2015, 2017]
low_years = [2008, 2016, 2019, 2020]
start_year = 2015

historical_data = load_historical_data(ticker)
phi = historical_data[PRICE_HISTORY]
dhi = historical_data[DIVIDEND_HISTORY]
ist = historical_data[INCOME_STATEMENT]

dphi = dhi.merge(phi, on='Date', how='right')
dphi.Amount = dphi.Amount.ffill().fillna(0)
dphi['Yield'] = dphi.Amount / dphi.Close
dphi = dphi[dphi.Yield != 0]
dphi = dphi[dphi.Date.dt.year > start_year]

highs = get_ovy(high_years, dhi, phi, ist)
lows = get_uvy(low_years, dhi, phi, ist)

uvy = (lambda x: sum(x) / len(x))([x['MIN_YLD'] for x in lows])
ovy = (lambda x: sum(x) / len(x))([x['MAX_YLD'] for x in highs])

dphi['Overvalued Yield'] = dphi.Amount / ovy
dphi['Undervalued Yield'] = dphi.Amount / uvy
dphi = dphi.loc[:, ['Date', 'Amount', 'Close',
                    'Yield', 'Overvalued Yield', 'Undervalued Yield']]

dphi['25% to OVV'] = dphi.loc[:, ['Overvalued Yield',
                                  'Undervalued Yield']].quantile(q=0.75, axis=1).values
dphi['25% to UVV'] = dphi.loc[:, ['Overvalued Yield',
                                  'Undervalued Yield']].quantile(q=0.25, axis=1).values

print(f'OVV = {ovy}, UVV={uvy}')
plotYieldChannels(dphi, ovy, uvy)
