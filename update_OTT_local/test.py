from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.http import HttpError

import gspread
from gspread_dataframe import get_as_dataframe, set_with_dataframe
from gspread_formatting.dataframe import format_with_dataframe, BasicFormatter
from gspread_formatting import *

import pandas as pd
from pandas._libs.tslibs.timestamps import Timestamp as tp

from datetime import date,datetime
import time
import xlsxwriter
import math
import numpy as np
import dateutil.parser

def ColNum2ColName(n):
	 convertString = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
	 base = 26
	 i = n - 1

	 if i < base:
			return convertString[i]
	 else:
			return ColNum2ColName(i//base) + convertString[i%base]

pd.set_option('mode.chained_assignment', None)
file_path = "/Users/sanjaykhadda/Downloads/pandas_simple.xlsx"
sheets_OTTData = ['Sheet1', 'Sheet2', 'Sheet3', 'Sheet4', 'Sheet5', 'Sheet6', 'Sheet7']
df_OTTData = []
for sheet in sheets_OTTData:
		df_OTTData.append(pd.read_excel(file_path, sheet_name=sheet, index=False))
j = 0
for df in df_OTTData:
		df = df.drop(df.columns[[0]], axis=1)
		for i in range(len(df.date)):
				try:
					df.date[i] = df.date[i].strftime("%m/%d/%y")
				except:
					print("j = ", j, " i = ", i)
		j = j+1



for i in range(len(df_OTTData[1].date)):
	df_OTTData[1]['keys[partner]'][i] = df_OTTData[1]['keys[partner]'][i].strip()
	if not df_OTTData[1]['keys[ad]'][i] or pd.isnull(df_OTTData[1]['keys[ad]'][i]):
		df_OTTData[1]['keys[ad]'][i] = "null"
	if not df_OTTData[1]['keys[partner]'][i] or pd.isnull(df_OTTData[1]['keys[partner]'][i]):
		df_OTTData[1]['keys[partner]'][i] = "null"

for i in range(len(df_OTTData[2].date)):
	df_OTTData[2]['keys[partner]'][i] = df_OTTData[2]['keys[partner]'][i].strip()
	if not df_OTTData[2]['keys[len]'][i] or pd.isnull(df_OTTData[2]['keys[len]'][i]):
		df_OTTData[2]['keys[len]'][i] = "null"
	if not df_OTTData[2]['keys[partner]'][i] or pd.isnull(df_OTTData[2]['keys[partner]'][i]):
		df_OTTData[2]['keys[partner]'][i] = "null"

for i in range(len(df_OTTData[3].date)):
	df_OTTData[3]['keys[partner]'][i] = df_OTTData[3]['keys[partner]'][i].strip()
	if not df_OTTData[3]['keys[provider]'][i] or pd.isnull(df_OTTData[3]['keys[provider]'][i]):
		df_OTTData[3]['keys[provider]'][i] = "null"

for i in range(len(df_OTTData[4].date)):
	df_OTTData[4]['keys[partner]'][i] = df_OTTData[4]['keys[partner]'][i].strip()
	if not df_OTTData[4]['keys[partner]'][i] or pd.isnull(df_OTTData[4]['keys[partner]'][i]):
		df_OTTData[4]['keys[partner]'][i] = "null"
	if not df_OTTData[4]['keys[network]'][i] or pd.isnull(df_OTTData[4]['keys[network]'][i]):
		df_OTTData[4]['keys[network]'][i] = "null"

for i in range(len(df_OTTData[5].date)):
	df_OTTData[5]['keys[partner]'][i] = df_OTTData[5]['keys[partner]'][i].strip()
	if not df_OTTData[5]['keys[partner]'][i] or pd.isnull(df_OTTData[5]['keys[partner]'][i]):
		df_OTTData[5]['keys[partner]'][i] = "null"
	if not df_OTTData[5]['keys[market]'][i] or pd.isnull(df_OTTData[5]['keys[market]'][i]):
		df_OTTData[5]['keys[market]'][i] = "null"

for i in range(len(df_OTTData[6].date)):
	df_OTTData[6]['keys[partner]'][i] = df_OTTData[6]['keys[partner]'][i].strip()
	if not df_OTTData[6]['keys[partner]'][i] or pd.isnull(df_OTTData[6]['keys[partner]'][i]):
		df_OTTData[6]['keys[partner]'][i] = "null"
	if not df_OTTData[6]['keys[devicetype]'][i] or pd.isnull(df_OTTData[6]['keys[devicetype]'][i]):
		df_OTTData[6]['keys[devicetype]'][i] = "null"


scope = ['https://spreadsheets.google.com/feeds',
				 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('/Users/sanjaykhadda/Desktop/update_OTT_local/client_secret.json', scope)
client = gspread.authorize(creds)

partners = ["CBSI", "Discovery", "PlutoTV", "Premion", "TubiTV", "hulu", "NBCU", "Viacom"]
sheet = client.open("Copy of RealReal - Alphonso - Task Tracker_1")

worksheets = []
partner_blocks = []
df_partner_blocks = []
for i in range(len(partners)):
		worksheets.append(sheet.worksheet("OTT Pixel Stats " + partners[i]))
		partner_blocks.append([])
		df_partner_blocks.append([])


for index in range(len(partners)):
		 start_block = 2
		 end_block = 6
		 for j in range(5):
				 block_j = []
				 for i in range(start_block, end_block):
						 try:
								 block_j.append(worksheets[index].col_values(i))
						 except Exception as e:
								 if str(e).find("RESOURCE_EXHAUSTED")>-1:
										 print("resource exhausted...came here with index = ", index, " j = ", j, " i = ", i)
										 done = False
										 while done is False:
												 try:
														 block_j.append(worksheets[index].col_values(i))
														 done = True
												 except:
															pass
								 else:
										 raise
				 start_block = end_block + 1
				 end_block = start_block + 4
				 partner_blocks[index].append(list(map(list, np.transpose(block_j)))[1:])

block_temp = [[],[]]
temp_index = 0
for index in [3,5]:
		 for i in range(27, 31):
				 try:
						 block_temp[temp_index].append(worksheets[index].col_values(i))
				 except Exception as e:
						 if str(e).find("RESOURCE_EXHAUSTED")>-1:
								 print("resource exhausted...came here with temp_index = ", temp_index, " i = ", i)
								 done = False
								 while done is False:
										 try:
												 block_temp[temp_index].append(worksheets[index].col_values(i))
												 done = True
										 except:
													pass
						 else:
								 raise
		 temp_index = temp_index + 1

partner_blocks[3].append(list(map(list, np.transpose(block_temp[0])))[1:])
partner_blocks[5].append(list(map(list, np.transpose(block_temp[1])))[1:])

date_from = tp(datetime.strptime(partner_blocks[0][0][0][1], "%m/%d/%y"))
today = tp(date.today().strftime("%m/%d/%y"))
latest_date = df_OTTData[0]['date'].max()
print("date_from = ", date_from, " today = ", today, "latest_date = ", latest_date)


index = 0
for partner in partners:
		num = 5
		for i in range(num):
				df_temp = df_OTTData[i].loc[df_OTTData[i]['date'] > date_from]
				df_temp = df_temp.loc[df_temp['date'] < latest_date]
				df_temp = df_temp[df_temp['keys[partner]'] == partner]
				print(df_temp)
				df_temp['date'] = pd.to_datetime(df_temp['date']).dt.date
				df_temp['date'] = df_temp['date'].transform(lambda x: str(x.strftime('%m/%d/%Y')))
				df_partner_blocks[index].append(df_temp)
		index = index + 1

df_temp = df_OTTData[5].loc[df_OTTData[5]['date'] > date_from]
df_temp = df_temp.loc[df_temp['date'] < latest_date]
df_temp = df_temp[df_temp['keys[partner]'] == "Premion"]
df_temp['date'] = pd.to_datetime(df_temp['date']).dt.date
df_temp['date'] = df_temp['date'].transform(lambda x: str(x.strftime('%m/%d/%Y')))
df_temp['keys[market]'] = df_temp['keys[market]'].transform(lambda x: str(int(x)) if x != "null" else x)
df_partner_blocks[3].append(df_temp)

df_temp = df_OTTData[6].loc[df_OTTData[6]['date'] > date_from]
df_temp = df_temp.loc[df_temp['date'] < latest_date]
df_temp = df_temp[df_temp['keys[partner]'] == "hulu"]
df_temp['date'] = pd.to_datetime(df_temp['date']).dt.date
df_temp['date'] = df_temp['date'].transform(lambda x: str(x.strftime('%m/%d/%Y')))
df_partner_blocks[5].append(df_temp)


for i in range(len(partners)):
		for j in range(len(df_partner_blocks[i])):
				df_partner_blocks[i][j] = df_partner_blocks[i][j].drop(['Unnamed: 0'], axis=1)

print("df = ", df_partner_blocks)

for i in range(len(partners)):
		df_partner_blocks[i][0] = pd.concat([df_partner_blocks[i][0],	pd.DataFrame(partner_blocks[i][0], columns = ['keys[partner]','date','param_cust','count'])])
		df_partner_blocks[i][1] = pd.concat([df_partner_blocks[i][1],	pd.DataFrame(partner_blocks[i][1], columns = ['keys[partner]','keys[ad]','date','count'])])
		df_partner_blocks[i][2] = pd.concat([df_partner_blocks[i][2],	pd.DataFrame(partner_blocks[i][2], columns = ['keys[partner]', 'keys[len]','date','count'])])
		df_partner_blocks[i][3] = pd.concat([df_partner_blocks[i][3],	pd.DataFrame(partner_blocks[i][3], columns = ['keys[partner]','keys[provider]','date','count'])])
		df_partner_blocks[i][4] = pd.concat([df_partner_blocks[i][4],	pd.DataFrame(partner_blocks[i][4], columns = ['keys[partner]','keys[network]','date','count'])])

df_partner_blocks[3][5] = pd.concat([df_partner_blocks[3][5],	pd.DataFrame(partner_blocks[3][5], columns = ['keys[partner]','keys[market]','date','count'])], sort=False)
df_partner_blocks[5][5] = pd.concat([df_partner_blocks[5][5],	pd.DataFrame(partner_blocks[5][5], columns = ['keys[partner]','date','keys[devicetype]','count'])], sort=False)

# print("df_ = ", df_partner_blocks)
index = 0
value_cell_format = cellFormat(
		horizontalAlignment='CENTER',
		borders = {
					"top": {
					"style": 'SOLID',
					},
					"bottom": {
						 "style": 'SOLID',
					},
					"left": {
						 "style": 'SOLID',
					},
					"right": {
						 "style": 'SOLID',
					}
		})

for worksheet in worksheets:
		column = 2

		for df in df_partner_blocks[index]:
				range = ColNum2ColName(column) + "1:" + ColNum2ColName(column+3) + str(len(df.index)+1)
				try:
						set_with_dataframe(worksheet, df, row=1, col=column)
						format_cell_range(worksheet, range, value_cell_format)
						column = column + 5
				except Exception as e:
						if str(e).find("RESOURCE_EXHAUSTED")>-1:
								print("resource exhausted...came here with index = ", index, " j = ", j, " i = ", i)
								done = False
								while done is False:
										try:
												set_with_dataframe(worksheet, df, row=1, col=column)
												format_cell_range(worksheet, range, value_cell_format)
												column = column + 5
												done = True
										except:
												 pass
						else:
								raise
		index = index + 1
