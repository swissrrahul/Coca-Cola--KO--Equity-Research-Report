# ┌────────────────────────────────────────────────────────────────────────────┐
# │                                 AXIOMEUS                                   │
# │                    Financial Statement Analysis Toolkit                    │
# └────────────────────────────────────────────────────────────────────────────┘
#
# ~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*
#
# Module      : Alpha Core
# Version     : 1.1
#
# Purpose:
#     Imports financial statement data directly from Excel
#     Automates ratio calculations that are traditionally performed manually
#     Calculates key financial ratios used in equity research and Common size statements
#     Formulates financial statement analysis.
#
# Features:
#     • Margin ratios
#     • Liquidity Ratios
#     • Solvency Ratios
#     • Profitability Ratios
#     • Efficiency Ratios
#     • Valuation Ratios
#     • Earnings Ratios
#     • Cashflow Ratios
#     • Return Ratios
#     • Others
#
# Author : Rahul Raj
#
# ~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*




#---------------------------**|| Importing Libraries ||**---------------------------#

import pandas as pd
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
pd.set_option('display.float_format', '{:.4f}'.format)



#---------------------------**|| Reading Excel File ||**---------------------------#

data_us = pd.read_excel("C:/Users/rahul/OneDrive/Documents/Equity Research Reports/The Coca-Cola Company/Ratiocalc_CocaCola.xlsx",
                        sheet_name="US_GAAP", usecols="B:H", header=2)
data_us.set_index("Metrics", inplace=True)

data_us_10yr = pd.read_excel("C:/Users/rahul/OneDrive/Documents/Equity Research Reports/The Coca-Cola Company/Ratiocalc_CocaCola.xlsx",
                        sheet_name="US_GAAP", usecols="B:N", header=2)
data_us_10yr.set_index("Metrics", inplace=True)




#---------------------------**|| Creating Functions ||**---------------------------#

def get_average(series):
    series_rev = series[::-1]
    avg = (series_rev + series_rev.shift(1)) / 2
    return avg[::-1]

def difference(series):
    series_rev = series[::-1]
    diff = (series_rev - series_rev.shift(1))
    return diff[::-1]

def divv(series):
    series_rev = series[::-1]
    diff = (series_rev / series_rev.shift(1))
    return diff[::-1]




#---------------------------**|| Global Variables ||**---------------------------#

os_shares = 4300723069

total_short_term_debt = data_us.loc["Current Maturities"] + data_us.loc["Loan"] + data_us.loc["Short Term Borrowings"]
invested_capital = ((data_us.loc["Total Assets"] - (data_us.loc["Total Cash & Bank Balance"] + data_us.loc["Marketable Securities"]
                    + data_us.loc["Short Term Investments"])) - (data_us.loc["Total Current Liabilities"] - total_short_term_debt))
working_capital = data_us.loc["Total Current Assets"] - data_us.loc["Total Current Liabilities"]
f_accounts_receivables = data_us.loc["Trade Receivable"] + data_us.loc["Accounts Receivable"]
f_accounts_payable = data_us.loc["Trade Payable"] + data_us.loc["Accounts Payable"]
normalized_ebt = data_us.loc["EBIT"] - data_us.loc["Total Other Income"]

average = pd.DataFrame({
"Avg Total Assets": get_average(data_us.loc["Total Assets"]),
"Avg Equity": get_average(data_us.loc["Total Equity"]),
"Avg Inventory": get_average(data_us.loc["Inventories"]),
"Avg Working Capital": get_average(working_capital),
"Avg Accounts Receivable": get_average(f_accounts_receivables),
"Avg Accounts Payable": get_average(f_accounts_payable)
})




#---------------------------**|| Creating Loops ||**---------------------------#

rolling_cagr = []
for i in range(0, 6):
    CAGR = (((data_us_10yr.loc["DEPS"].iloc[i]) / (data_us_10yr.loc["DEPS"].iloc[i + 5])) ** (1/5)) - 1
    rolling_cagr.append(CAGR)
rolling_cagr = (pd.Series(rolling_cagr, index=data_us.columns) * 100)
print(rolling_cagr)



#---------------------------**|| Header ||**---------------------------#

print("""
◆━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━◆
                       ♦ ALPHA CORE ♦
◆━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━◆
""")




#---------------------------**|| Margins ||**---------------------------#

margins = pd.DataFrame({
    "Gross Margin": data_us.loc["Gross Profit"] / data_us.loc["Net Revenue"],
    "EBITDA Margin": data_us.loc["EBITDA"] / data_us.loc["Net Revenue"],
    "EBIT Margin": data_us.loc["EBIT"] / data_us.loc["Net Revenue"],
    "EBT Margin": data_us.loc["EBT"] / data_us.loc["Net Revenue"],
    "Net Margin": data_us.loc["Net Income"] / data_us.loc["Net Revenue"]
    })




#---------------------------**|| Ratios from Balance sheet ||**---------------------------#

liquidity_ratios = pd.DataFrame({
    "Current Ratio": data_us.loc["Total Current Assets"] / data_us.loc["Total Current Liabilities"],
    "Quick Ratio": (data_us.loc["Total Current Assets"]-data_us.loc["Inventories"])/data_us.loc["Total Current Liabilities"],
    "Cash Ratio": (data_us.loc["Total Cash & Bank Balance"] + data_us.loc["Marketable Securities"] + data_us.loc["Short Term Investments"])
                  / data_us.loc["Total Current Liabilities"]
    })


earnings_ratio = pd.DataFrame({
    "Interest Coverage": data_us.loc["EBIT"] / data_us.loc["Interest exp"],
    "Effective Tax Rate": data_us.loc["Tax"] / data_us.loc["EBT"]
    })


solvency_ratios = pd.DataFrame({
    "Debt to Equity Ratio":(data_us.loc["Loan"]+data_us.loc["Long Term Debt"]+data_us.loc["Long Term Borrowings"]
    + data_us.loc["Current Maturities"]) /data_us.loc["Total Equity"],
    "Debt to EBITDA": (data_us.loc["Loan"] + data_us.loc["Long Term Debt"] + data_us.loc["Long Term Borrowings"]
    + data_us.loc["Current Maturities"]) / data_us.loc["EBITDA"],
    "Debt Ratio": (data_us.loc["Loan"]+data_us.loc["Long Term Debt"]+data_us.loc["Long Term Borrowings"]
    + data_us.loc["Current Maturities"]) / data_us.loc["Total Assets"],
    "Equity Multiplier": data_us.loc["Total Assets"] / data_us.loc["Total Equity"],
    "Interest Coverage": data_us.loc["EBIT"] / data_us.loc["Interest exp"]
    })

valuation_ratios = pd.DataFrame({
    "Price to Earnings": data_us.loc["Stock Price"] / data_us.loc["DEPS"]
    })
valuation_ratios["Price to Earnings Growth"] = valuation_ratios["Price to Earnings"] / rolling_cagr


efficiency_activity_ratios = pd.DataFrame({
    "Receivable Turnover": data_us.loc["Net Revenue"] / average["Avg Accounts Receivable"],
    "Inventory Turnover": data_us.loc["COGS"] / average["Avg Inventory"],
    "Total Asset Turnover": data_us.loc["Net Revenue"] / average["Avg Total Assets"],
    "Working Capital Turnover": data_us.loc["Net Revenue"] / average["Avg Working Capital"]
    })


others = pd.DataFrame({
    "Change in Working Capital": difference(working_capital),
    "DIO": 365 / efficiency_activity_ratios["Inventory Turnover"],
    "DSO": 365 / efficiency_activity_ratios["Receivable Turnover"],
    "DPO": 365 / (data_us.loc["COGS"] / average["Avg Accounts Payable"]),
    })
others["CCC"] = others["DIO"] + others["DSO"] - others["DPO"]


nopat = data_us.loc["EBIT"] * (1 - earnings_ratio["Effective Tax Rate"])

fcff = (
    nopat
    + data_us.loc["Depreciation & Amortization in CFS"]
    - data_us.loc["CapEx"]
    - others["Change in Working Capital"]
    )


fcf_metrics = pd.DataFrame({
    "FCF": fcff,
    "FCF Margin": fcff / data_us.loc["Net Revenue"]
    })
fcf_metrics["FCF Growth"] = difference(fcf_metrics["FCF"])/(fcf_metrics["FCF"].shift(-1))




#---------------------------**|| Normalized FCFF ||**---------------------------#

normalized_effective_tax = data_us.loc["Tax"] / normalized_ebt
normalized_nopat = data_us.loc["EBIT"] * (1 - normalized_effective_tax)

normalized_fcff = (
    normalized_nopat
    + data_us.loc["Depreciation & Amortization in CFS"]
    - data_us.loc["CapEx"]
    - others["Change in Working Capital"]
    )
normalized_fcf = pd.DataFrame({"Normalized FCF": normalized_fcff,
                               "Normalized FCF Margin": fcff / data_us.loc["Net Revenue"]})
normalized_fcf["Normalized FCF Growth"] = difference(normalized_fcf["Normalized FCF"])/(normalized_fcf["Normalized FCF"].shift(-1))


return_ratios = pd.DataFrame({
    "Return on Assets": data_us.loc["Net Income"] / average["Avg Total Assets"],
    "Return on Equity": data_us.loc["Net Income"] / average["Avg Equity"],
    "Return on Invested Capital": normalized_nopat / invested_capital
    })


cashflow_ratios = pd.DataFrame({
    "Quality of Earnings": data_us.loc["Net Cashflow from Operating Activities"] / data_us.loc["Net Income"],
    "FCF Ratio": normalized_fcff / data_us.loc["EBITDA"],
    "CFO to Revenue": data_us.loc["Net Cashflow from Operating Activities"] / data_us.loc["Net Revenue"]
    })




#---------------------------**|| Transposing Rows & Columns ||**---------------------------#

margins = margins.T
liquidity_ratios = liquidity_ratios.T
solvency_ratios = solvency_ratios.T
return_ratios = return_ratios.T
earnings_ratio = earnings_ratio.T
efficiency_activity_ratios = efficiency_activity_ratios.T
fcf_metrics = fcf_metrics.T
normalized_fcf = normalized_fcf.T
others = others.T
valuation_ratios = valuation_ratios.T
cashflow_ratios = cashflow_ratios.T




#---------------------------**|| Print Statements ||**---------------------------#

print("\n(i) Profit Margins\n",margins)
print("\n\n(ii) Liquidity Ratios\n",liquidity_ratios)
print("\n\n(iii) Solvency/Leverage Ratios\n",solvency_ratios)
print("\n\n(iv) Return Ratios\n",return_ratios)
print("\n\n(v) Earnings Ratio\n",earnings_ratio)
print("\n\n(vi) Valuation Ratios\n",valuation_ratios)
print("\n\n(vii) Efficiency/Activity Ratios\n",efficiency_activity_ratios)
print("\n\n(viii) Cashflow Ratios\n", cashflow_ratios)
print("\n\n(viii) FCFF Metrics\n",fcf_metrics,"\n\n~Reported free cash flow is derived from cash flow from operations less capital expenditures. "
                  "\nThis measure reflects the company’s total cash generation, including the impact of non-operating items, "
                  "\nfinancing-related flows, and period-to-period working capital movements~")
print("\n\n(ix) Normalized FCFF Metrics\n",normalized_fcf,"\n\n~For valuation purposes, free cash flow has been normalized to reflect the sustainable cash-generating"
                     "\ncapacity of the company’s core operations. Adjustments have been made to exclude non-operating income, "
                     "\ninvestment-related earnings, and other non-recurring or non-core items~")
print("\n\n(x) Others\n",others)












