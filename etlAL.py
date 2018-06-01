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

#------ program starting point --------	
if __name__=="__main__":
	#--------- database info ----------------
	driver = r"{Microsoft Access Driver (*.mdb, *.accdb)};"

	db_file = r"C:\\pycode\\DirAnnualBenefits\\RDDDBenefits.accdb;"
	user = "admin"
	password = ""
	#--------------------------------------------------------------------

	table = "tbl_AchLevel"
	columns = '''
	[Cslt],
	[FirstName],
	[SurName],
	[Director],
	[Area],
	[Region],
	[RO],
	[AppointmentDate],
	[TenureinIG],
	[Status],
	[StatusEffectiveDate],
	[TenureinRole],
	[BirthDate],
	[Age],
	[Language],
	[MDPoints],
	[TOPPoints],
	[AL],
	[ALMix],
	[InRole],
	[LANID],
	[CycleDate]
	'''
	
	dfal = pd.read_excel('C:\\pycode\\DirAnnualBenefits\\2017 data to set 2018 Dir AL.xlsx')
	dfal['CycleDate'] = '12/31/2017'
	
	print dfal.dtypes
	print dfal.head()
	print dfal.tail()

	dbq.add_to_tbl(driver, db_file, table, columns, dfal)

	print 'done'
	