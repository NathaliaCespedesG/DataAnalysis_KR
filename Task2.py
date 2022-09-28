import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import PatternFill
#from itertools import chain

def reading_files():

	name_file1 = '1MC03-SCJ-IM-REP-SS01_SL03-290400_P11_Data.xlsx'

	all_dfs1 = pd.read_excel(name_file1, sheet_name = None)

	#Converting into one tab 

	all_dfs1.keys() # Esto leer los nombres de los tabs

	df_1 = pd.concat(all_dfs1, ignore_index = True)

	# Combining the files (new and old) into one
	writer = pd.ExcelWriter('Filter.xlsx', engine='xlsxwriter')
	df_1.to_excel(writer, sheet_name = 'All design packages')

	writer.save()

def separating_tabs():

	name_file = 'Filter.xlsx'

	writer = pd.ExcelWriter('Filter.xlsx', engine='xlsxwriter')
	df = pd.read_excel(name_file, sheet_name = 'All design packages')
	df.drop(columns = df.columns[0], axis = 1, inplace= True)
	ECS_MW = df[df['115 Design Package'] == 'ECS-MW']
	ECS_CC = df[df['115 Design Package'] == 'ECS-CC']
	ECS_CW = df[df['115 Design Package'] == 'ECS-CW']
	blanks = df[df['115 Design Package'].isnull()]
	#cond = df['115 Design Package'].isna()
	#blanks = df[df['115 Design Package'] == 'NaN']

	ECS_MW.to_excel(writer, sheet_name = 'ECS-MW')
	ECS_CC.to_excel(writer, sheet_name = 'ECS-CC')
	ECS_CW.to_excel(writer, sheet_name = 'ECS-CW')
	blanks.to_excel(writer, sheet_name = 'blanks')


	writer.save()


m = reading_files()
n = separating_tabs()