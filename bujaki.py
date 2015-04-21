from math import ceil
import urllib2

from bs4 import BeautifulSoup
import xlsxwriter

TSX60 = ('AEM', 'AGU', 'ATD.B', 'ARX', 'ABX', 'BCE', 'BB', 'BBD.B',
         'CCO', 'CNR', 'CNQ', 'COS', 'CP', 'CTC.A', 'CCT', 'CVE',
         'GIB.A', 'CPG', 'ENB', 'ECA', 'FTS', 'WN', 'GIL', 'G',
         'HSE', 'IMO', 'K', 'L', 'MG', 'MRU', 'POT',
         'RCI.B', 'SAP', 'SJR.B', 'SLW', 'SNC', 'SU', 'TLM', 'TCK.B',
         'T', 'TRI', 'TA', 'TRP', 'VRX', 'YRI', 'TD', 'BMO',
         'BNS', 'RY', 'BAM.A') #, 'MFC', 'POW', 'SLF', 'CF')

BASE_URL = "http://web.tmxmoney.com/financials.php?qm_symbol=%s&type=BalanceSheet&&rtype=A"
BASE_URL1 = "http://web.tmxmoney.com/financials.php?qm_symbol=%s&type=IncomeStatement&rtype=A"

workbook = xlsxwriter.Workbook('All.xlsx')
worksheet = workbook.add_worksheet()
col = 0

has_first_col = False

for company in TSX60[30:]:
    request = urllib2.urlopen(BASE_URL%(company))
    request1 = urllib2.urlopen(BASE_URL1%(company))
    html_file = request.read()
    html_file1 = request1.read()
    soup = BeautifulSoup(html_file)
    soup1 = BeautifulSoup(html_file1)
    balance_sheet = soup.findAll("table")[1]
    income_statement = soup1.findAll("table")[1]
    years = balance_sheet.findAll("thead")[0]
    years1 = income_statement.findAll("thead")[0]
    year = years.findAll("th")[2]
    year1 = years1.findAll("th")[2]
    elements = {}
    elements1 = {}
    indexes = []
    indexes1 = []
    rows = balance_sheet.findAll("tr")[1:]
    rows1 = income_statement.findAll("tr")[1:]

    for row in rows:
        item = row.findAll("td")[0].text.strip()
        indexes.append(item)

    for row in rows1:
        item = row.findAll("td")[0].text.strip()
        indexes1.append(item)

    for i, row in enumerate(rows):
        value = int(row.findAll("td")[2].text.replace(",", "0").replace("--", "0").replace("(", "-").replace(")", ""))
        elements[indexes[i]] = value

    for i, row in enumerate(rows1):
        value = row.findAll("td")[2].text.replace(",", "").replace("--", "0").replace("(", "-").replace(")", "")
        elements1[indexes1[i]] = value


    TA = elements['Total Assets']
    def converter(component):
        return '{:.2%}'.format(float(component)/TA)

    def ratios (num, denom):
        a = int(num)
        b = int(denom)
        ratio = '{:.2%}'.format(float(a)/float(b))
        return ratio

    from collections import OrderedDict
    data = OrderedDict()
    data["Cash"] = converter(elements['Cash Cash Equivalents And Short Term Investments']if "Cash Cash Equivalents And Short Term Investments" in elements else (elements['Cash Cash Equivalents And Federal Funds Sold'])) #else elements['Short Term Investments'])
    data["Accounts Receivalbe"] = converter(elements['Receivables'] if "Receivables" in elements else (elements["Net Loans"] + elements["Receivables"]) )
    data["Inventory"] = converter(elements['Inventory'] if "Inventory" in elements else "0")
    data["Other current assets"] = converter(elements['Other Current Assets'] if 'Other Current Assets' in elements else elements['Cash Cash Equivalents And Federal Funds Sold'] + elements['Receivables'] - elements['Cash Cash Equivalents And Federal Funds Sold'] - elements["Net Loans"] - elements["Receivables"] - elements['Inventory'] if "Inventory" in elements else "0")
    data["Total current assets"] = converter(elements['Current Assets'] if "Current Assets" in elements else (elements['Cash Cash Equivalents And Federal Funds Sold'] + elements['Receivables']))
    data["Plant & equipment"] = converter(elements['Net PPE'])
    data["Goodwill & equipment"] = converter(elements['Goodwill'])
    data["Other intangible assets"] = converter(elements['Other Intangible Assets']) 
    data["Other noncurrent assets"] = converter(elements['Other Non Current Assets'] if 'Other Current Assets' in elements else elements['Total Assets'] - elements['Cash Cash Equivalents And Federal Funds Sold'] - elements['Receivables'] - elements['Net PPE'] - elements['Other Intangible Assets'] - elements['Goodwill'])
    data["Total noncurrent assets"] = '{:.2%}'.format(float(TA - elements['Total Non Current Assets'] if 'Total Non Current Assets' in elements else (elements['Net PPE'] + elements['Goodwill And Other Intangible Assets'] + elements['Separate Account Assets'] + elements['Other Assets']))/TA)
    data["Accounts payable"] = converter(elements['Payables And Accrued Expenses'])
    data["Prepaid assets"] = converter(elements["Prepaid Assets"])
    data["Other current liabilities"] = converter(elements['Other Current Liabilities'] if 'Other Current Liabilities' in elements else (elements['Current Deferred Liabilities']))
    data["Total current liabilities"] = converter(elements['Current Liabilities'] if 'Current Liabilities' in elements else (elements['Current Deferred Liabilities'] + elements['Payables And Accrued Expenses']))
    data["Long term debt"] = converter(elements['Long Term Debt And Capital Lease Obligation'])
    data["Long term provisions"] = converter((elements['Long Term Provisions']))
    data["Capital lease obligations"] = converter((elements['Capital Lease Obligations']))
    data["Other non-current liabilities"] = converter(elements['Other Non Current Liabilities'] if 'Other Non Current Liabilities' in elements else (elements['Non Current Deferred Liabilities'] + elements['Non Current Accrued Expenses']))
    data["Minority interest"] = converter(elements['Minority Interest'])
    data["Total liabilities"] = converter(elements['Total Liabilities'])
    data["Capital stock"] = converter(elements['Capital Stock'])
    data["Retained earnings"] = converter(elements['Retained Earnings'])
    data["Other equity"] = converter(elements['Gains Losses Not Affecting Retained Earnings'])
    data["Total stockholders' equity"] = converter(elements['Stockholders Equity'])
    data["Gross margin"] = ratios(elements1['Gross Profit'] if 'Gross Profit' in elements1 else elements1['Interest Income After Provision For Loan Loss'], elements1['Total Revenue'])
    data["R&D/sales"] = ratios(elements1['Research And Development'], elements1['Total Revenue'])
    data["Profit margin"] = ratios(elements1['Net Income'], elements1['Total Revenue'])
    data["Day of receivables"] = round((float(elements['Receivables'] if "Receivables" in elements else (elements["Net Loans"] + elements["Receivables"])*365))/(float(elements1['Total Revenue'])), 2)
    data['Inventory turnover'] = 0 
    inventory = float(elements['Inventory'] if "Inventory" in elements else "0")
    if inventory > 0 and "Cost Of Revenue" in elements1:
        data['Inventory turnover'] = round(float(elements1['Cost Of Revenue'])/inventory, 2)
    data["Fixed asset turnover"] = ratios(elements1['Total Revenue'], elements['Net PPE'])
    data["Total Asset Turnover"] = ratios(elements1['Total Revenue'], elements['Total Assets'])
    data["Return on assets"] = ratios(elements1['Net Income'], elements['Total Assets'])
    data["Return on equity"] = ratios(elements1['Net Income'], elements['Stockholders Equity'])
    data["Debt to equity"] = ratios(elements['Long Term Debt And Capital Lease Obligation'], elements['Stockholders Equity'])


#    for keys,values in data.items():
#        print(keys)
#        print(values)

    print company
    
    row = 1
    worksheet.write(1, col+1, company)

    for key, value in data.items():
        row += 1
        if not has_first_col:
            worksheet.write(row, col, key)
        worksheet.write(row, col+1, value)

    if not has_first_col:
        has_first_col = True

    col += 1

workbook.close()
    
    

