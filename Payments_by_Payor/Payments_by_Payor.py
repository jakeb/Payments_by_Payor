#!/usr/bin/env python
# coding: utf-8

# Destination: C:\Users\BridgeJ\Box\PFS Projects\Reports\Managed Care\
# 
# Schedule: Weekly, on Mondays

# In[1]:


import pandas as pd
import math

from datetime import datetime, date, timedelta
import os
import shutil
import pyodbc
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo


def connect(server, db):
    driver = [x for x in pyodbc.drivers() if 'SQL Server' in x][-1]
    cnx = pyodbc.connect(
        f"Driver={driver};"
        f"Server={server};"
        f"Database={db};"
        "Trusted_Connection=yes;"
    )
    return cnx   


# In[2]:


def df_to_excel(
    df,
    excel_path,
    sheet_name="sheet1",
    money_cols=[],
    perc_cols=[],
    num_cols=[],
    text_cols=[],
    date_cols=[],
    outindex=False,
    close_file=True
):
    """
    Exports a dataframe to an excel file with common formatting.
    Fields with names ending in '_amount', '_amt', '$' get formatted as currency.
    Fields with names ending in '_perc', '%' get formatted as percentage.
    Fields with names ending in 'date' get date formatting 
    
    Should I look at column types as well as names?
    """
    # write dataframe to excel
    if type(excel_path) is pd.io.excel._XlsxWriter:
        outfile = excel_path
    else:
        outfile = pd.ExcelWriter(excel_path)

    df.to_excel(outfile, index=outindex, sheet_name=sheet_name)

    if len(df.index) > 0:
        # basic formatting
        workbook = outfile.book
        perc_fmt = workbook.add_format({"num_format": "0.0%", "bold": False})
        money_fmt = workbook.add_format({"num_format": "$#,##0.00", "bold": False})
        count_fmt = workbook.add_format({"num_format": "0", "bold": False})
        date_fmt = workbook.add_format({"num_format": "d/m/yyyy"})
        text_fmt = workbook.add_format({"num_format": "@"})

        ws = outfile.sheets[sheet_name]
        col = 0
        col_lengths = [
            min(100, max(len(c), math.floor(df[c].astype(str).map(len).max() * 1.25)))
            for c in df
        ]

        for a in list(df):
            if (a[-1] == "$") or (a in money_cols):
                ws.set_column(col, col, max(10, col_lengths[col]), money_fmt)
            elif (a[-1] == "#") or (a in num_cols):
                ws.set_column(col, col, max(7, col_lengths[col]), count_fmt)
            elif (a[-1] == "%") or (a in perc_cols):
                ws.set_column(col, col, max(10, col_lengths[col]), perc_fmt)
            elif a in text_cols:
                ws.set_column(col, col, max(10, col_lengths[col]), text_fmt)
            elif a in date_cols:
                ws.set_column(col, col, max(10, col_lengths[col]), date_fmt)
            else:
                ws.set_column(col, col, col_lengths[col])
            col += 1

    # outfile.save()
    if close_file:
        outfile.close()
        return excel_path
    else:
        return outfile


# In[3]:


[server, db] = ['apexclarityprd.ucsfmedicalcenter.org', 'Clarity']
cnx = connect(server, db)


# In[4]:


# date setup

start_date = '20200401'  

det_start_date = '20200601'
# end date = previous Sunday, start date for detail report is previous Monday


# In[5]:



 
today = date.today()
last_monday = today - timedelta(days=today.weekday())
previous_monday = today + timedelta(days=-today.weekday(), weeks=-1)
print("Today:", today)
print("Last Monday:", last_monday)
print("Previous Monday:", previous_monday)


# In[6]:




setup_q = f"""


IF OBJECT_ID('tempdb..#results') IS NOT NULL
    DROP TABLE #results

IF OBJECT_ID('tempdb..#summary') IS NOT NULL
    DROP TABLE #summary

IF OBJECT_ID('tempdb..#optum_united') IS NOT NULL
    DROP TABLE #optum_united



IF OBJECT_ID('tempdb..#payor_groups') IS NOT NULL
    DROP TABLE #payor_groups
	 
CREATE TABLE #payor_groups (
	payor_id int not null primary key,
	payor varchar(30) null,
	product varchar(50) null
	)

INSERT INTO #payor_groups (payor_id, payor, product) 
	Values
	(124,'Aetna ','Commercial/Managed Care Plans'),
	(125,'Aetna ','Medicare Plans'),
	(126,'Aetna ','Medicare Plans'),
	(138,'Anthem','Commercial/Managed Care Plans'),
	(140,'Anthem','Medicare Plans'),
	(141,'Anthem','Medicaid/MediCal Plans'),
	(142,'Anthem','Medicare Plans'),
	(156,'Blue Shield','Commercial/Managed Care Plans'),
	(157,'Blue Shield ','Medicare Plans'),
	(181,'Cigna','Commercial/Managed Care Plans'),
	(236,'Health Net','Medicare Plans'),
	(237,'Health Net ','Medicaid/MediCal Plans'),
	(372,'United','Commercial/Managed Care Plans'),
	(374,'United ','Commercial/Managed Care Plans'),
	(375,'United ','Medicare Plans'),
	(376,'United ','Commercial/Managed Care Plans'),
	(377,'United ','Medicare Plans'),
	(404,'Aetna ','Commercial/Managed Care Plans'),
	(406,'Anthem','Commercial/Managed Care Plans'),
	(407,'Anthem','Commercial/Managed Care Plans'),
	(408,'Blue Shield','Commercial/Managed Care Plans'),
	(411,'Cigna','Commercial/Managed Care Plans'),
	(412,'Health Net','Commercial/Managed Care Plans'),
	(425,'United/Optum ','Transplant'),
	(457,'Anthem','Commercial/Managed Care Plans'),
	(458,'Blue Shield','Commercial/Managed Care Plans'),
	(461,'Health Net','Commercial/Managed Care Plans'),
	(467,'Anthem','Commercial/Managed Care Plans'),
	(468,'Blue Shield','Commercial/Managed Care Plans'),
	(479,'Anthem','Medicaid/MediCal Plans'),
	(481,'Health Net ','Medicaid/MediCal Plans'),
	(529,'Anthem','UC Care'),
	(560,'Aetna ','Medicaid/MediCal Plans'),
	(562,'United','Medicaid/MediCal Plans'),
	(587,'Aetna ','Medicaid/MediCal Plans'),
	(591,'Blue Shield ','Medicaid/MediCal Plans'),
	(492,'Aetna','Commercial/Managed Care Plans'),
	(235,'Health Net','Commercial/Managed Care Plans')



Select 
	convert(date, har.ADM_DATE_TIME ) as admit_date,
	convert(date, har.DISCH_DATE_TIME) as disch_date,
    har.TOT_ACCT_BAL, bkt.CURRENT_BALANCE , 
    bs.NAME as [Bill Status], bks.name as [Bucket Status] ,

	case when har.loc_id = 20000 then 'BCHO' else 'UCSF' end	as [Location]	,
	case when har.loc_id = 20000 then '94-0382330' else '94-3281657' end	Tax_ID,
    p.payor,
	p.product,
	--htr.tx_id,
	htr.bucket_id,
	htr.hsp_account_id,
	htr.payor_id,
	htr.SERVICE_DATE ,
	htr.TX_POST_DATE ,
	htr.tx_amount,
	htr.INVOICE_NUM ,
	htr.INT_CONTROL_NUMBER

into #results
from 
	hsp_account har
    join HSP_TRANSACTIONS htr on htr.HSP_ACCOUNT_ID = har.HSP_ACCOUNT_ID 
	join #payor_groups p on p.payor_id = htr.PAYOR_ID 
	join HSP_BUCKET bkt on bkt.BUCKET_ID = htr.BUCKET_ID 
	join ZC_ACCT_BILLSTS_HA bs on har.ACCT_BILLSTS_HA_C = bs.ACCT_BILLSTS_HA_C 
	join ZC_BKT_STS_HA bks on bks.BKT_STS_HA_C = bkt.BKT_STS_HA_C 

where 
	htr.TX_POST_DATE >= '{start_date}' 

	and htr.TX_TYPE_HA_C =2 and htr.tx_amount <> 0 
	and htr.SERV_AREA_ID = 10


Select distinct b.bucket_id
into #optum_united
 from
 #results b
 join hsp_account har on har.HSP_ACCOUNT_ID = b.hsp_account_id and b.payor_id = 425
 join patient pat on pat.pat_id = har.pat_id
 join hsp_account har2 on pat.pat_id = har2.pat_id and har.HSP_ACCOUNT_ID <> har2.hsp_account_id 
 join hsp_bucket bk on bk.HSP_ACCOUNT_ID = har2.HSP_ACCOUNT_ID 
 join coverage cvg on cvg.coverage_id = bk.coverage_id
 join COVERAGE_MEMBER_LIST cvg_mem on cvg_mem.COVERAGE_ID = cvg.COVERAGE_ID 

 where cvg.PAYOR_ID in
( 
372, 
562,
374,
375,
376,
377   -- ORIGINAL UNITED EPM LIST

) 

DELETE FROM #results where bucket_id in (Select bucket_id from #optum_united)



"""
cursor = cnx.cursor()

cursor.execute(setup_q)
cursor.commit()


# In[7]:


detail = pd.read_sql(f"Select * from #results where tx_post_date >= '{previous_monday}'", cnx)
#detail


# In[8]:


def summary_q(loc):
    return f"""
    Select
        r.payor, r.product,
        -sum(CASE WHEN month(r.tx_post_date) = 4 then  r.tx_amount end) as [Apr-20],
        -sum(CASE WHEN month(r.tx_post_date) = 5 then  r.tx_amount end) as [May-20],
        -sum(CASE WHEN month(r.tx_post_date) = 6 then  r.tx_amount end) as [Jun-20],
        -sum(CASE WHEN month(r.tx_post_date) = 7 then r.tx_amount end) as [Jul-20]


    from #results r


        where r.[location] = '{loc}'

    group by 

            r.payor, r.product
            ORDER BY r.payor, r.product
    """


# In[9]:


ucsf_sum = pd.read_sql(summary_q('UCSF'), cnx)
bcho_sum = pd.read_sql(summary_q('BCHO'), cnx)


# In[10]:


filename = r'C:\Users\BridgeJ\Box\PFS Projects\Reports\Managed Care' + f"\\payments_by_payor APeX {last_monday}.xlsx"
outf = df_to_excel(detail, filename,sheet_name='detail', close_file=False)
df_to_excel(ucsf_sum, outf,sheet_name='UCSF summary', close_file=False)
df_to_excel(bcho_sum,outf,sheet_name='BCHO summary', close_file=False)
outf.close()


# ## format results as table

# In[11]:


print(filename)


# In[12]:


wb = load_workbook(filename)
tab_reference = f"A1:{chr(64+detail.shape[1])}{detail.shape[0]+1}"
tab = Table(displayName="payment_detail", ref= tab_reference)
ws = wb['detail']
# I list out the 4 show-xyz options here for reference
style = TableStyleInfo(
    name="TableStyleLight18",    #light green
    showFirstColumn=False,
    showLastColumn=False,
    showRowStripes=True,
    showColumnStripes=False
)
tab.tableStyleInfo = style
ws.add_table(tab)

if 1==0:
    ws = wb['UCSF summary']
    tab_reference = f"A1:{chr(64+ucsf_sum.shape[1])}{ucsf_sum.shape[0]+1}"
    tab = Table(displayName="ucsf_summary", ref= tab_reference)
    tab.tableStyleInfo = style
    ws.add_table(tab)

    ws = wb['BCHO summary']
    tab_reference = f"A1:{chr(64+bcho_sum.shape[1])}{bcho_sum.shape[0]+1}"
    tab = Table(displayName="bcho_summary", ref= tab_reference)
    tab.tableStyleInfo = style
    ws.add_table(tab)


# In[13]:


wb.save(filename)


# In[14]:


wb.close()


# In[ ]:




