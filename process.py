import os, re, sys, time, xlrd, pyodbc, datetime
from datetime import date
import fnmatch
import numpy as np
import pandas as pd
import itertools as it
from openpyxl import load_workbook
from shutil import copyfile
import myfun as dd
import dbquery as dbq

def cal_tenure(pncycle):
	tenure = round(float(pncycle)/24, 2)
	return tenure

def try_float(pnum):
    try:
        floatnum = float(pnum)
    except ValueError:
        return pnum
    else:
        return floatnum

def split_str_to_col (pstr):
	#strip left blanks and end '\n', '\r'
	ltstr = filter(None, re.split(r'\s{2,}', pstr.lstrip(' ').rstrip('\n').rstrip('\r')))

	for idx in range(len(ltstr)):
		#remove front and trailing blank for each element and remove ',' seperator for numbers
		ltstr[idx] = ltstr[idx].lstrip(' ').rstrip(' ').replace(',', '')
		#update '-' to front to show correct negitive amount
		if '-' in ltstr[idx]:
			try:
				float(ltstr[idx].replace('-', ''))
			except ValueError:
				ltstr[idx]
			else:
				ltstr[idx] = float('-' + ltstr[idx].replace('-', ''))
		ltstr[idx] = try_float(ltstr[idx])	
	#ltstr = [try_float(x) for x in ltstr]		
	return ltstr

def transfer_txt_to_ds (pstr):	
	dfoutputdata = []	
	cycledate = ''
	outputnamedate = ''
	outputname = '' #use for accumulator num
	
	with open(pstr) as f:
		for line in f:
			if line.strip():
				#print 'in strip'
				# get cycle end date
				if CycDatePa.match(line.lstrip(' ')) and len(cycledate) == 0:
					#get start date and end date to a list, set cycledate to end date
					#cycledate = re.search('20\d{2}\s+\D{3}\s+\d{1,2}', line).group()
					cycledate = re.findall('20\d{2}\s+\D{3}\s+\d{1,2}', line)[1]
					#print cycledate
					tempd = datetime.datetime.strptime(cycledate, '%Y %b %d')
					cycledate = tempd.strftime('%m/%d/%Y')
					outputnamedate = tempd.strftime('%Y%m%d')
					#break
				# get ACCUMULATOR TYPE 
				if accumtyppattern.match(line[2:].lstrip(' ')) and len(outputname) == 0:
					outputname = re.search(r'\d+', line[2:]).group()			
					#break
				#get normal data lines	
				if DataPa.match(line[2:].lstrip(' ')):
					dfoutputdata.append(split_str_to_col(line[2:]))
				#get totals from file
				if re.match(r'TOTAL ACCUMULATED AMOUNT', line[2:].lstrip(' ')):
					filetotal = split_str_to_col(line[2:])
					#print filetotal
						
	#print 'before assign'
	labels = ['CNSLT NUM', 'CACT TYPE', 'CURRENT DEALERSHIP', 'IGFS ACCUMULATED AMOUNT', 'IGSI ACCUMULATED AMOUNT', 'TOTAL ACCUMULATED AMOUNT']
	df = pd.DataFrame(dfoutputdata, columns=labels)
		
	df['CYCLE END DATE'] = cycledate#datetime.datetime.strptime(cycledate, '%m/%d/%Y')

	print 'now handle ' + outputname
	#print df.dtypes
	if np.isclose(df['IGSI ACCUMULATED AMOUNT'].sum(), float(filetotal[2])):
		print 'IGFI ACCUMULATED AMOUNT matches' 
	if np.isclose(df['TOTAL ACCUMULATED AMOUNT'].sum(), filetotal[3]):
		print 'TOTAL ACCUMULATED AMOUNT matches' 	

	return df

	
	
#------ program starting point --------	
if __name__=="__main__":
	## dd/mm/yyyy format
	print 'Process date is ' + str(time.strftime("%m/%d/%Y"))
	print 'Please enter the cycle end date (mm/dd/yyyy) you want to process:'
	#-----------------------------------------------------
	#------- get cycle date ----------------------
	getcycledate = datetime.datetime.strptime(raw_input(), '%m/%d/%Y')
	endday = getcycledate
	startday = datetime.datetime.strptime('1/1/' + str(endday.year), '%m/%d/%Y')

	print 'Cycle start date is ' + str(startday)
	print 'Cycle end date is ' + str(endday)
	#--------- database info ----------------
	driver = r"{Microsoft Access Driver (*.mdb, *.accdb)};"

	db_file = r"C:\\pycode\\DirAnnualBenefits\\RDDDBenefits.accdb;"
	user = "admin"
	password = ""
	#--------------------------------------------------------------------
	#--------- get DD list and No of cycle they served for the year ----------
	sql = '''
			SELECT 
				qry_ActiveDD.LKG_CSLT_NUM AS [DD]
				,qry_ActiveDD.LKG_CSLT_NAM_FULL AS [FullName]
				,qry_ActiveDD.[CStatus]
				,qry_ActiveDD.[CTermDate]
				,qry_ActiveDD.[CTermReason]
				,qry_ActiveDD.[CPosition]
				,qry_ActiveDD.LKG_CSLT_LANGUAGE AS [Language]
				,MIN(qry_ActiveDD.LKG_CSLT_SMPL_DTE) AS [StartDate]
				,MAX(qry_ActiveDD.LKG_CSLT_SMPL_DTE) AS [EndDate]
				,COUNT(qry_ActiveDD.LKG_CSLT_SMPL_DTE) AS [NofCycle]
			FROM qry_ActiveDD
			WHERE (qry_ActiveDD.LKG_CSLT_SMPL_DTE) BETWEEN #''' + startday.strftime("%m/%d/%Y") + '''# AND #''' + endday.strftime("%m/%d/%Y") + '''#
			GROUP BY 
				qry_ActiveDD.LKG_CSLT_NUM
				,qry_ActiveDD.LKG_CSLT_NAM_FULL
				,qry_ActiveDD.[CStatus]
				,qry_ActiveDD.[CTermDate]
				,qry_ActiveDD.[CTermReason]
				,qry_ActiveDD.[CPosition]
				,qry_ActiveDD.LKG_CSLT_LANGUAGE
			ORDER BY 	
				qry_ActiveDD.LKG_CSLT_NUM
		'''

	dfdd = dbq.df_select(driver, db_file, sql)

	dds = ','.join(['%s' % x for x in dfdd.loc[dfdd['EndDate'] != endday.strftime("%m/%d/%Y"), 'DD']]) #get all cslts left the role during the year

	sql = '''
			SELECT 
				qry_CsltTerm.LKG_CSLT_NUM AS [DD]
				,MAX(qry_CsltTerm.LKG_CSLT_SMPL_DTE) AS [EndDate]
			FROM qry_CsltTerm
			WHERE ((qry_CsltTerm.LKG_CSLT_SMPL_DTE) BETWEEN #''' + startday.strftime("%m/%d/%Y") + '''# AND #''' + endday.strftime("%m/%d/%Y") + '''#)
				AND qry_CsltTerm.LKG_CSLT_NUM IN (%s)
			GROUP BY 
				qry_CsltTerm.LKG_CSLT_NUM
		'''	
	sql = sql % (dds)
	
#-------------------------------------	
#	#print sql
#	with open("Output.txt", "w") as text_file:
#		text_file.write(sql)
#	sys.exit("done")	
#--------------------------------------	
	dftermddreason = pd.DataFrame()
	dftermdd = dbq.df_select(driver, db_file, sql)
	
	for index, row in dftermdd.iterrows():
		sql = '''
			SELECT 
				qry_CsltTerm.LKG_CSLT_NUM AS [DD]
				,qry_CsltTerm.LKG_CSLT_STATUS
				,qry_CsltTerm.LKG_CSLT_TERM_DTE
				,qry_CsltTerm.LKG_CSLT_TERM_REASON1
				,qry_CsltTerm.LKG_CSLT_TERM_REASON2
			FROM qry_CsltTerm
			WHERE (qry_CsltTerm.LKG_CSLT_SMPL_DTE) = #''' + row['EndDate'].strftime("%m/%d/%Y") + '''#
				AND qry_CsltTerm.LKG_CSLT_NUM = ''' + str(row['DD']) + ''';
		'''	
		dftermddreason = dftermddreason.append(dbq.df_select(driver, db_file, sql), ignore_index=True)
	
	dfdd = dfdd.merge(dftermddreason, on='DD', how='left')
	dfdd['LKG_CSLT_STATUS'].fillna('Active', inplace=True)

	#-----------------------------------------------------------------------

	#--------- get RD list and No of cycle they served for the year ----------
	sql = '''
			SELECT 
				qry_ActiveRD.LKG_CSLT_NUM AS [RD]
				,qry_ActiveRD.LKG_CSLT_NAM_FULL AS [FullName]
				,qry_ActiveRD.[CStatus]
				,qry_ActiveRD.[CTermDate]
				,qry_ActiveRD.[CTermReason]
				,qry_ActiveRD.[CPosition]
				,qry_ActiveRD.LKG_CSLT_LANGUAGE AS [Language]
				,MIN(qry_ActiveRD.LKG_CSLT_SMPL_DTE) AS [StartDate]
				,MAX(qry_ActiveRD.LKG_CSLT_SMPL_DTE) AS [EndDate]
				,COUNT(qry_ActiveRD.LKG_CSLT_SMPL_DTE) AS [NofCycle]
			FROM qry_ActiveRD
			WHERE (qry_ActiveRD.LKG_CSLT_SMPL_DTE) BETWEEN #''' + startday.strftime("%m/%d/%Y") + '''# AND #''' + endday.strftime("%m/%d/%Y") + '''#
			GROUP BY 
				qry_ActiveRD.LKG_CSLT_NUM
				,qry_ActiveRD.LKG_CSLT_NAM_FULL
				,qry_ActiveRD.[CStatus]
				,qry_ActiveRD.[CTermDate]
				,qry_ActiveRD.[CTermReason]
				,qry_ActiveRD.[CPosition]
				,qry_ActiveRD.LKG_CSLT_LANGUAGE
			ORDER BY 	
				qry_ActiveRD.LKG_CSLT_NUM
		'''

	dfrd = dbq.df_select(driver, db_file, sql)

	rds = ','.join(['%s' % x for x in dfrd.loc[dfrd['EndDate'] != endday.strftime("%m/%d/%Y"), 'RD']]) #get all cslts left the role during the year

	sql = '''
			SELECT 
				qry_CsltTerm.LKG_CSLT_NUM AS [RD]
				,MAX(qry_CsltTerm.LKG_CSLT_SMPL_DTE) AS [EndDate]
			FROM qry_CsltTerm
			WHERE ((qry_CsltTerm.LKG_CSLT_SMPL_DTE) BETWEEN #''' + startday.strftime("%m/%d/%Y") + '''# AND #''' + endday.strftime("%m/%d/%Y") + '''#)
				AND qry_CsltTerm.LKG_CSLT_NUM IN (%s)
			GROUP BY 
				qry_CsltTerm.LKG_CSLT_NUM
		'''	
	sql = sql % (rds)
	
#-------------------------------------	
#	#print sql
#	with open("Output.txt", "w") as text_file:
#		text_file.write(sql)
#	sys.exit("done")	
#--------------------------------------	
	dftermrdreason = pd.DataFrame()
	dftermrd = dbq.df_select(driver, db_file, sql)
	
	for index, row in dftermrd.iterrows():
		sql = '''
			SELECT 
				qry_CsltTerm.LKG_CSLT_NUM AS [RD]
				,qry_CsltTerm.LKG_CSLT_STATUS
				,qry_CsltTerm.LKG_CSLT_TERM_DTE
				,qry_CsltTerm.LKG_CSLT_TERM_REASON1
				,qry_CsltTerm.LKG_CSLT_TERM_REASON2
			FROM qry_CsltTerm
			WHERE (qry_CsltTerm.LKG_CSLT_SMPL_DTE) = #''' + row['EndDate'].strftime("%m/%d/%Y") + '''#
				AND qry_CsltTerm.LKG_CSLT_NUM = ''' + str(row['RD']) + ''';
		'''	
		dftermrdreason = dftermrdreason.append(dbq.df_select(driver, db_file, sql), ignore_index=True)
	
	dfrd = dfrd.merge(dftermrdreason, on='RD', how='left')
	dfrd['LKG_CSLT_STATUS'].fillna('Active', inplace=True)
	#print dfrd.head()
	#sys.exit('The process is stopped')
	#-----------------------------------------------------------------------	
	
	#--------- handle ACCUMULATOR to get business income -------------------
	#DataPa = re.compile(r'^\d{1,5}\s{2,}\d{1}\s{2,}IG\D{2}\s{1}\(\d{4}\).*$')
	DataPa = re.compile(r'\d{1,5}\s{2,}\d{1}\s{2,}\D{4}.*$')
	accumtyppattern = re.compile(r'ACCUMULATOR TYPE.*\d+')
	CycDatePa = re.compile(r'.*THRU\s+20\d{2}\s+\D{3}\s+\d{1,2}')
	Negative = re.compile(r'-')

	filelist = [] 

	for file in os.listdir('C:\\pycode\\DirAnnualBenefits\\'):
		if file.endswith('.txt'):
			filelist.append(os.path.join('C:\\pycode\\DirAnnualBenefits\\', file))

	#print filelist	
	for txts in filelist:
		with open(txts) as f:
			for line in f:
				if line.strip():
					if accumtyppattern.match(line[2:].lstrip(' ')):
						accumulatortype = re.search(r'\d+', line[2:]).group()
						print 'Will start to process ACCUMULATOR TYPE ' + accumulatortype
						if accumulatortype == '1257':
							dfoutput = transfer_txt_to_ds(txts)
							dfytdddbi = dfoutput.groupby(['CNSLT NUM'], as_index=False)['TOTAL ACCUMULATED AMOUNT'].sum()
						elif accumulatortype == '1259':
							dfoutput = transfer_txt_to_ds(txts)
							dfytdrdbi = dfoutput.groupby(['CNSLT NUM'], as_index=False)['TOTAL ACCUMULATED AMOUNT'].sum()
						elif accumulatortype == '881':
							dfoutput = transfer_txt_to_ds(txts)
							dfytdrdai = dfoutput.groupby(['CNSLT NUM'], as_index=False)['TOTAL ACCUMULATED AMOUNT'].sum()	
						break
	#----------------------------------------------------------------------

	#---------- DD part ----------------------------------
	#---------- Get DD Business Income ------------------------
	dfdd = dfdd.merge(dfytdddbi, left_on='DD', right_on='CNSLT NUM', how='left') #get DD business income from 1257
	
	dfr12 = pd.read_excel('C:\\pycode\\DirAnnualBenefits\\R12.xlsx')
	dfdd = dfdd.merge(dfr12[['REP NUM', 'YTD DD BUS INC']], left_on='DD', right_on='REP NUM', how='left') #get DD business income from R12
	dfdd['Diff'] = dfdd['TOTAL ACCUMULATED AMOUNT'].fillna(0).round(2) - dfdd['YTD DD BUS INC'].fillna(0).round(2) #compare 1257 vs R12 number

	#--------- database info ----------------
	driver = r"{Microsoft Access Driver (*.mdb, *.accdb)};"

	db_file = r"C:\\pycode\\DirAnnualBenefits\\RDDDBenefits.accdb;"
	user = "admin"
	password = ""
	#--------------------------------------------------------------------

	#--------- get DD tenure before last year ----------
	sql = '''
			SELECT DISTINCT
				[tbl_Tenure].[Cslt] AS [DD]
				,[tbl_Tenure].[Tenure]
			FROM [tbl_Tenure]
			WHERE [tbl_Tenure].[Position] = 'DD' AND ([tbl_Tenure].[Period]) = #''' + endday.strftime("%m/%d/%Y") + '''#
		'''

	dftenure = dbq.df_select(driver, db_file, sql)
	#dftenure = pd.read_excel('C:\\pycode\\DirAnnualBenefits\\Tenure.xlsx', sheet_name='DD')
	dfdd = dfdd.merge(dftenure, on='DD', how='left') #get previous tenure
	dfdd['CYearTenure'] = dfdd['NofCycle'].apply(cal_tenure) #get current year tenure
	dfdd['FinalTenure'] = dfdd['Tenure'].fillna(0) + dfdd['CYearTenure'] #get real tenure
	dfdd['TenureLevel'] = np.where(dfdd['FinalTenure'] < 5, 1, 0)
	dfdd['TenureLevel'] = np.where((dfdd['FinalTenure'] >= 5) & (dfdd['FinalTenure'] <= 10), 2, dfdd['TenureLevel'])
	dfdd['TenureLevel'] = np.where(dfdd['FinalTenure'] > 10, 3, dfdd['TenureLevel'])

	#--------- get DD AL ----------
	sql = '''
			SELECT DISTINCT
				[tbl_AchLevel].[Cslt] AS [DD]
				,[tbl_AchLevel].[AL]
			FROM [tbl_AchLevel]
			WHERE [tbl_AchLevel].[Status] = 'DD' AND ([tbl_AchLevel].[Period]) = #''' + endday.strftime("%m/%d/%Y") + '''#
		'''

	dfal = dbq.df_select(driver, db_file, sql)
	dfdd = dfdd.merge(dfal, on='DD', how='left') #get AL 
	#for index, row in dfdd[dfdd['AL'].isnull()].iterrows():
	#	if row['StartDate'] >= datetime.datetime.strptime('1/15/' + str(startday.year), '%m/%d/%Y'):
	#		row['AL'] = 4
	#print dfdd[['DD','AL','StartDate']]
	#sys.exit('---------stop---------')
	#dfdd['AL'].fillna(0)

	#--------- get DD Deferred Income ----------
	sql = '''
			SELECT DISTINCT
				[DDDeferredIncome].[Tenure] AS [TenureLevel]
				,[DDDeferredIncome].[AchievementLevel] AS [AL]
				,[DDDeferredIncome].[Rate] AS [DIRate]
			FROM [DDDeferredIncome]
			WHERE ([DDDeferredIncome].[Period]) = #''' + endday.strftime("%m/%d/%Y") + '''#
		'''

	dfdi = dbq.df_select(driver, db_file, sql)
	dfdd = dfdd.merge(dfdi, on=['AL','TenureLevel'], how='left') #get DD Deferred Income rate
	#dfdd['DIRate'].fillna(0)
	dfdd.loc[dfdd['CTermReason'].str.contains("COMPETITOR|CONFORMING", na=False), 'DIRate'] = np.NAN #If terminated to the competitor or for not confirming then they do not have anything calculated
	dfdd.loc[dfdd['LKG_CSLT_TERM_REASON1'].str.contains("COMPETITOR|CONFORMING", na=False), 'DIRate'] = np.NAN #If terminated to the competitor or for not confirming then they do not have anything calculated
 	dfdd['DeferredIncome'] = dfdd['TOTAL ACCUMULATED AMOUNT'] * dfdd['DIRate']
	dfdd['DIInstallment'] = dfdd['DeferredIncome']/5

	#--------- get DD Benefit Credit ----------
	sql = '''
			SELECT DISTINCT
				[DDBenefitCredit].[AchievementLevel] AS [AL]
				,[DDBenefitCredit].[LumpSum] AS [BCLumpSum]
				,[DDBenefitCredit].[Rate] AS [BCRate]
			FROM [DDBenefitCredit]
			WHERE ([DDBenefitCredit].[Period]) = #''' + endday.strftime("%m/%d/%Y") + '''#
		'''

	dfbc = dbq.df_select(driver, db_file, sql)
	dfdd = dfdd.merge(dfbc, on='AL', how='left') #get DD Benefit Credit rate
	#dfdd['BCLumpSum'].fillna(0)
	#dfdd['BCRate'].fillna(0)
	dfdd.loc[dfdd['CTermReason'].str.contains("COMPETITOR|CONFORMING", na=False), ['BCLumpSum', 'BCRate']] = np.NAN #If terminated to the competitor or for not confirming then they do not have anything calculated
	dfdd.loc[dfdd['LKG_CSLT_TERM_REASON1'].str.contains("COMPETITOR|CONFORMING", na=False), ['BCLumpSum', 'BCRate']] = np.NAN #If terminated to the competitor or for not confirming then they do not have anything calculated
 	dfdd.loc[dfdd['CStatus'] != 'Active', ['BCLumpSum', 'BCRate']] = np.NAN #if terminated, no benefit credits or education credits
	dfdd['BCLumpSumAmt'] = dfdd['BCLumpSum'] * dfdd['CYearTenure']
	dfdd['BCRateAmt'] = dfdd['TOTAL ACCUMULATED AMOUNT'] * dfdd['BCRate']
	dfdd['BenefitCredit'] = dfdd['BCLumpSumAmt'] + dfdd['BCRateAmt']

	#--------- get DD Education Credit ----------
	sql = '''
			SELECT DISTINCT
				[DDEducationCredit].[Quartile] AS [AL]
				,[DDEducationCredit].[Rate] AS [ECRate]
			FROM [DDEducationCredit]
			WHERE ([DDEducationCredit].[Period]) = #''' + endday.strftime("%m/%d/%Y") + '''#
		'''

	dfbc = dbq.df_select(driver, db_file, sql)
	dfdd = dfdd.merge(dfbc, on='AL', how='left') #get DD Education Credit
	#dfdd['ECRate'].fillna(0)
	dfdd.loc[dfdd['CTermReason'].str.contains("COMPETITOR|CONFORMING", na=False), 'ECRate'] = np.NAN
	dfdd.loc[dfdd['LKG_CSLT_TERM_REASON1'].str.contains("COMPETITOR|CONFORMING", na=False), 'ECRate'] = np.NAN
	dfdd.loc[dfdd['CStatus'] != 'Active', 'ECRate'] = np.NAN #if terminated, no benefit credits or education credits	
	dfdd['EducationCredit'] = dfdd['ECRate'] * dfdd['CYearTenure']

	dfdd.drop(['CNSLT NUM','REP NUM','TenureLevel','CPosition'], axis=1, inplace=True)
	dfdd.rename(columns={'LKG_CSLT_STATUS':'Status', 'TOTAL ACCUMULATED AMOUNT':'BusinessIncome', 'YTD DD BUS INC':'R12BI', 'Tenure':'LYearTenure'}, inplace=True)
	
	#---------- RD part ----------------------------------
	#---------- Get RD Business Income ------------------------
	dfrd = dfrd.merge(dfytdrdbi, left_on='RD', right_on='CNSLT NUM', how='left') #get RD business income from 1259
	dfrd.rename(columns={'TOTAL ACCUMULATED AMOUNT':'BI'}, inplace=True)
	dfrd.drop(['CNSLT NUM'], axis=1, inplace=True)
	dfrd = dfrd.merge(dfytdrdai, left_on='RD', right_on='CNSLT NUM', how='left') #get RD asset income from 881
	dfrd.rename(columns={'TOTAL ACCUMULATED AMOUNT':'AI'}, inplace=True)
	dfrd.drop(['CNSLT NUM'], axis=1, inplace=True)
	dfrd['BusinessIncome'] = dfrd['BI'].fillna(0).round(2) + dfrd['AI'].fillna(0).round(2)
		
	dfrd = dfrd.merge(dfr12[['REP NUM', 'YTD RD BUS INC']], left_on='RD', right_on='REP NUM', how='left') #get RD business income from R12
	dfrd['Diff'] = dfrd['BusinessIncome'].fillna(0).round(2) - dfrd['YTD RD BUS INC'].fillna(0).round(2) #compare business income vs R12 number

	#--------- database info ----------------
	driver = r"{Microsoft Access Driver (*.mdb, *.accdb)};"

	db_file = r"C:\\pycode\\DirAnnualBenefits\\RDDDBenefits.accdb;"
	user = "admin"
	password = ""
	#--------------------------------------------------------------------

	#--------- get RD tenure ----------
	sql = '''
			SELECT DISTINCT
				[tbl_Tenure].[Cslt] AS [RD]
				,[tbl_Tenure].[Tenure]
			FROM [tbl_Tenure]
			WHERE [tbl_Tenure].[Position] = 'RD' AND ([tbl_Tenure].[Period]) = #''' + endday.strftime("%m/%d/%Y") + '''#
		'''

	dftenure = dbq.df_select(driver, db_file, sql)
	#dftenure = pd.read_excel('C:\\pycode\\DirAnnualBenefits\\Tenure.xlsx', sheet_name='DD')
	dfrd = dfrd.merge(dftenure, on='RD', how='left') #get previous tenure
	dfrd['CYearTenure'] = dfrd['NofCycle'].apply(cal_tenure) #get current year tenure
	dfrd['FinalTenure'] = dfrd['Tenure'].fillna(0) + dfrd['CYearTenure'] #get real tenure
	dfrd['TenureLevel'] = np.where(dfrd['FinalTenure'] < 5, 1, 0)
	dfrd['TenureLevel'] = np.where((dfrd['FinalTenure'] >= 5) & (dfrd['FinalTenure'] <= 10), 2, dfrd['TenureLevel'])
	dfrd['TenureLevel'] = np.where((dfrd['FinalTenure'] > 10) & (dfrd['FinalTenure'] <= 15), 3, dfrd['TenureLevel'])
	dfrd['TenureLevel'] = np.where(dfrd['FinalTenure'] > 15, 4, dfrd['TenureLevel'])

	#--------- get RD AL ----------
	sql = '''
			SELECT DISTINCT
				[tbl_AchLevel].[Cslt] AS [RD]
				,[tbl_AchLevel].[AL]
			FROM [tbl_AchLevel]
			WHERE [tbl_AchLevel].[Status] = 'RD' AND ([tbl_AchLevel].[Period]) = #''' + endday.strftime("%m/%d/%Y") + '''#
		'''

	dfal = dbq.df_select(driver, db_file, sql)
	dfrd = dfrd.merge(dfal, on='RD', how='left') #get AL 
	#dfrd['AL'].fillna(0)

	#--------- get RD Deferred Income ----------
	sql = '''
			SELECT DISTINCT
				[RDDeferredIncome].[Tenure] AS [TenureLevel]
				,[RDDeferredIncome].[AchievementLevel] AS [AL]
				,[RDDeferredIncome].[Rate] AS [DIRate]
			FROM [RDDeferredIncome]
			WHERE ([RDDeferredIncome].[Period]) = #''' + endday.strftime("%m/%d/%Y") + '''#
		'''

	dfdi = dbq.df_select(driver, db_file, sql)
	dfrd = dfrd.merge(dfdi, on=['AL','TenureLevel'], how='left') #get RD Deferred Income rate
	#dfrd['DIRate'].fillna(0)
	dfrd.loc[dfrd['CTermReason'].str.contains("COMPETITOR|CONFORMING", na=False), 'DIRate'] = np.NAN #If terminated to the competitor or for not confirming then they do not have anything calculated
	dfrd.loc[dfrd['LKG_CSLT_TERM_REASON1'].str.contains("COMPETITOR|CONFORMING", na=False), 'DIRate'] = np.NAN #If terminated to the competitor or for not confirming then they do not have anything calculated
 	dfrd['DeferredIncome'] = dfrd['BusinessIncome'] * dfrd['DIRate']
	dfrd['RegionBuilderDI'] = np.NAN #preserve for Region Builder Award Benefit Credits 
	dfrd['TotalDI'] = dfrd['DeferredIncome'] + dfrd['RegionBuilderDI'].fillna(0)
	dfrd['TotalDIInstallment'] = dfrd['TotalDI']/5
	
	#--------- get RD Benefit Credit ----------
	sql = '''
			SELECT DISTINCT
				[RDBenefitCredit].[AchievementLevel] AS [AL]
				,[RDBenefitCredit].[LumpSum] AS [BCLumpSum]
				,[RDBenefitCredit].[Rate] AS [BCRate]
			FROM [RDBenefitCredit]
			WHERE ([RDBenefitCredit].[Period]) = #''' + endday.strftime("%m/%d/%Y") + '''#
		'''

	dfbc = dbq.df_select(driver, db_file, sql)
	dfrd = dfrd.merge(dfbc, on='AL', how='left') #get RD Benefit Credit rate
	#dfrd['BCLumpSum'].fillna(0)
	#dfrd['BCRate'].fillna(0)
	dfrd.loc[dfrd['CTermReason'].str.contains("COMPETITOR|CONFORMING", na=False), ['BCLumpSum', 'BCRate']] = np.NAN #If terminated to the competitor or for not confirming then they do not have anything calculated
	dfrd.loc[dfrd['LKG_CSLT_TERM_REASON1'].str.contains("COMPETITOR|CONFORMING", na=False), ['BCLumpSum', 'BCRate']] = np.NAN #If terminated to the competitor or for not confirming then they do not have anything calculated
 	dfrd.loc[dfrd['CStatus'] != 'Active', ['BCLumpSum', 'BCRate']] = np.NAN #if terminated, no benefit credits or education credits
 	dfrd['BCLumpSumAmt'] = dfrd['BCLumpSum'] * dfrd['CYearTenure']
	dfrd['BCRateAmt'] = dfrd['BusinessIncome'] * dfdd['BCRate']
	dfrd['RegionBuilderBC'] = np.NAN #preserve for Region Builder Award Deferred Income
	dfrd['TotalBC'] = dfrd['BCLumpSumAmt'] + dfrd['BCRateAmt'] + dfrd['RegionBuilderBC'].fillna(0)

	#--------- get RD Education Credit ----------
	sql = '''
			SELECT DISTINCT
				[RDEducationCredit].[Quartile] AS [AL]
				,[RDEducationCredit].[Rate] AS [ECRate]
			FROM [RDEducationCredit]
			WHERE ([RDEducationCredit].[Period]) = #''' + endday.strftime("%m/%d/%Y") + '''#
		'''

	dfbc = dbq.df_select(driver, db_file, sql)
	dfrd = dfrd.merge(dfbc, on='AL', how='left') #get RD Education Credit
	dfrd.loc[dfrd['CTermReason'].str.contains("COMPETITOR|CONFORMING", na=False), 'ECRate'] = np.NAN
	dfrd.loc[dfrd['LKG_CSLT_TERM_REASON1'].str.contains("COMPETITOR|CONFORMING", na=False), 'ECRate'] = np.NAN
	dfrd.loc[dfrd['CStatus'] != 'Active', 'ECRate'] = np.NAN #if terminated, no benefit credits or education credits	
	dfrd['EducationCredit'] = dfrd['ECRate'] * dfrd['CYearTenure']
	
	#--------- get RD Admin Support Bonus ----------
	sql = '''
			SELECT DISTINCT
				[RDAdminSupportBonus].[Quartile] AS [AL]
				,[RDAdminSupportBonus].[Rate] AS [ASBonus]
			FROM [RDAdminSupportBonus]
			WHERE ([RDAdminSupportBonus].[Period]) = #''' + endday.strftime("%m/%d/%Y") + '''#
		'''

	dfasb = dbq.df_select(driver, db_file, sql)
	dfrd = dfrd.merge(dfasb, on='AL', how='left') #get RD Admin Support Bonus
	dfrd.loc[dfrd['CStatus'] != 'Active', 'ASBonus'] = np.NAN #if terminated, no AS Bouns
	dfrd.loc[dfrd['CPosition'] != 'REGIONAL DIRECTOR', 'ASBonus'] = np.NAN #if not in RD role, no AS Bouns
	#print dfrd.head()
	#sys.exit('--------stop---------')
	
	dfrd.drop(['REP NUM','TenureLevel','CPosition'], axis=1, inplace=True)
	dfrd.rename(columns={'LKG_CSLT_STATUS':'Status', 'YTD RD BUS INC':'R12BI', 'Tenure':'LYearTenure'}, inplace=True)
	
	#--------- output to Excel ---------------------
	writer = pd.ExcelWriter('DD.xlsx', engine='xlsxwriter')
	
	dfdd.to_excel(writer, sheet_name='DD', startrow=1, freeze_panes=(2,2), index=False)
	dfrd.to_excel(writer, sheet_name='RD', startrow=1, freeze_panes=(2,2), index=False)
	workbook = writer.book
	worksheet = writer.sheets['DD']
	worksheet2 = writer.sheets['RD']
	
	# Add some cell formats.
	formatcell = workbook.add_format({'bold': True, 'align':'center'})
	formatcslt = workbook.add_format({'bold':True, 'bg_color':'#FFFF00'})
	formatdate = workbook.add_format({'num_format':'mm/dd/yyyy'})
	formatnum = workbook.add_format({'num_format':'#,##0.00'})
	formatpercent = workbook.add_format({'num_format':'0.0000%'})
	formatbi = workbook.add_format({'num_format':'#,##0.00', 'bold':True, 'bg_color':'#FFFF00'})

	#----------- add DD notes --------------
	worksheet.write('P1', 'N-O', formatcell)
	worksheet.write('S1', 'Q+R', formatcell)
	worksheet.write('V1', 'N*U', formatcell)
	worksheet.write('W1', 'V/5', formatcell)
	worksheet.write('Z1', 'X*R', formatcell)
	worksheet.write('AA1', 'N*Y', formatcell)
	worksheet.write('AB1', 'Z+AA', formatcell)
	
	#----------- add RD notes --------------	
	worksheet2.write('R1', 'P-Q', formatcell)
	worksheet2.write('U1', 'S+T', formatcell)
	worksheet2.write('X1', 'P*W', formatcell)
	worksheet2.write('Z1', 'X+Y', formatcell)
	worksheet2.write('AA1', 'Z/5', formatcell)
	worksheet2.write('AD1', 'AB*T', formatcell)	
	worksheet2.write('AE1', 'P*AC', formatcell)
	worksheet2.write('AG1', 'AD+AE+AF', formatcell)
	
	
	# Set the column width and format for DD tab
	worksheet.set_column('A:A', 8, formatcslt)
	worksheet.set_column('B:B', 15)
	worksheet.set_column('D:D', 10, formatdate)
	worksheet.set_column('G:H', 10, formatdate)
	worksheet.set_column('N:N', 10, formatbi)
	worksheet.set_column('O:S', 10, formatnum)
	worksheet.set_column('T:T', 5, formatcslt)
	worksheet.set_column('U:U', 10, formatpercent)
	worksheet.set_column('V:X', 14, formatnum)
	worksheet.set_column('Y:Y', 10, formatpercent)
	worksheet.set_column('Z:AD', 14, formatnum)

	# Set the column width and format for RD tab
	worksheet2.set_column('A:A', 8, formatcslt)
	worksheet2.set_column('B:B', 15)
	worksheet2.set_column('D:D', 10, formatdate)
	worksheet2.set_column('G:H', 10, formatdate)
	worksheet2.set_column('N:O', 10, formatnum)
	worksheet2.set_column('P:P', 10, formatbi)
	worksheet2.set_column('Q:U', 10, formatnum)
	worksheet2.set_column('V:V', 5, formatcslt)
	worksheet2.set_column('W:W', 10, formatpercent)
	worksheet2.set_column('X:AB', 14, formatnum)
	worksheet2.set_column('AC:AC', 10, formatpercent)
	worksheet2.set_column('AD:AK', 12, formatnum)	
	
	# Close the Pandas Excel writer and output the Excel file.
	writer.save()
	
	print 'The process is done'
	
	
	