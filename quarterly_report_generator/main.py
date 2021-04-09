import os
import xlrd
import xlwt 
from xlwt import Workbook 
from openpyxl import load_workbook
from datetime import datetime


def write_2d_list(sheet, row, col, twodlist, ind):
	style = xlwt.XFStyle()

	# font
	font = xlwt.Font()
	font.bold = True
	style.font = font
	columnwidth = {}
	rowx = 1
	for rowdata in twodlist:
	    column = col
	    for colomndata in rowdata:
	        if column in columnwidth:
	            columnwidth[column] = max(len(colomndata), len("Conversion Rate"), columnwidth[column])

	        else:
	            columnwidth[column] = len(colomndata)
	        column = column + 1
	    rowx = rowx + 1	 
	for r in range(len(twodlist)):
		for c in range(len(twodlist[0])):
			if ind == 0 and r == 0:
				sheet.write(row + r, col + c, twodlist[r][c], style=style)
			else:
				sheet.write(row + r, col + c, twodlist[r][c])
	for column, widthvalue in columnwidth.items():
		sheet.col(column).width = (widthvalue + 2) * 256
	for r in range(1, 10000):
		sheet.row(r).height_mismatch = True
		sheet.row(r).height = 310


def search_string(sheet, string):
	locs = []
	for rowidx in range(sheet.nrows):
		row = sheet.row(rowidx)
		for colidx, cell in enumerate(row):
			if str(cell.value).strip() == string:
				locs.append((colidx, rowidx))
	return locs

def find_rows_cols(cur_sheet, row, col):
	col_end = col
	row_end = row
	while True:
		try:
			if cur_sheet.cell_value(row, col_end).strip() == "":
				break
			else:
				pass
		except IndexError:
			break					
		col_end = col_end + 1						
	while True:
		try:
			if cur_sheet.cell_value(row_end, col).strip() == "":
				break
			else:
				pass
		except IndexError:
			break
		row_end = row_end + 1
	return row_end, col_end

def main():
	months = {
		"Jan" : "01", 
		"Feb" : "02", 
		"Mar" : "03", 
		"Apr" : "04", 
		"May" : "05", 
		"Jun" : "06", 
		"Jul" : "07", 
		"Aug" : "08", 
		"Sep" : "09", 
		"Oct" : "10", 
		"Nov" : "11", 
		"Dec" : "12", 
	}
	Year = "2020"
	sheet_list = ['TV-Network', 'TV-Daypart', 'TV-Show', 'TV-Creative', 'TV-Content Duration', 'TV-Day Of the Week', 'TV-Creative Daypart', 'TV-Day of the week Daypart', 'TV-Network Daypart', 'TV-NW-DP-Len', 'TV-Network Show', 'TV-Network Day of the Week', 'TV-Recency', 'TV-Frequency', 'OTT-Partner', 'OTT-Partner Provider', 'OTT-Partner Market', 'OTT-Partner Network', 'OTT-Partner Daypart', 'OTT-Partner Day of the Week', 'OTT-Partner Campaign', 'OTT-Daypart', 'OTT-Day Of the Week', 'OTT-Frequency', 'OTT-Recency']


	reports = []
	quarterly_data = []

	for file in os.listdir("./reports"):
		if file.endswith(".xlsx"):
		    file_name = file.split("-")
		    # print(file)
		    start_date = months[file_name[1]] + "/" + file_name[2] + "/" + Year
		    end_date = months[file_name[3]] + "/" + file_name[4] + "/" + Year
		    week = months[file_name[1]] + "/" + file_name[2] + " - " + months[file_name[3]] + "/" + file_name[4]
		    reports.append({"month": file_name[1], "week": week, "start_date": datetime.strptime(start_date, "%m/%d/%Y"), "end_date": datetime.strptime(end_date, "%m/%d/%Y"), "report" : xlrd.open_workbook("./reports/" + file), "name": file})


	reports.sort(key = lambda x: x["start_date"], reverse = True)
	# print(reports[0]["report"].sheet_names())
	# for report in reports:
	# 	print(report["name"])

	for sheet in sheet_list:
		sheet_data = {
		"Registration": [], 
		"Purchase": [],
		"Purchase-new": [],
		"Purchase-repeat": []
		}
		first_report = {
		"Registration": True, 
		"Purchase": True,
		"Purchase-new": True,
		"Purchase-repeat": True
		}
		for report in reports:
			print(report["name"])
			try:
				cur_sheet = report["report"].sheet_by_name(sheet)
				evt_coords = search_string(cur_sheet, "Event Name")
				for evt in evt_coords:
					print(evt)
					print(first_report)
					col, row = evt
					event = cur_sheet.cell_value(row + 1, col)
					event_data = []
					row_end, col_end = find_rows_cols(cur_sheet, row, col)
					if first_report[event] == False:
						row = row + 1
						print("event1 = ", event)
					for r in range(row, row_end):
						row_data = []
						if r == row and first_report[event] == True:
							row_data.append("Month")
							row_data.append("Week of")
						else:
							row_data.append(report["month"])
							row_data.append(report["week"])
						for c in range(col, col_end):
							row_data.append(str(cur_sheet.cell_value(r, c)).strip())
						event_data.append(row_data)
					
					# print("event = ", event)
					sheet_data[event].append(event_data)
					first_report[event] = False
			except:
				pass
				# print("event_data = ", event_data)
		quarterly_data.append(sheet_data)		
	
	# print(quarterly_data[3]["Registration"])
	wb = Workbook() 
	
	# borders = xlwt.Borders()
	# borders.bottom = xlwt.Borders.DASHED
	# style.borders = borders
	title = "Alphonso TV Attribution Insights - The RealReal - TRR_Jan_06_Apr_05"
	for i in range(len(sheet_list)):
		worksheet = wb.add_sheet(sheet_list[i])
		start_row = 4
		start_col = 0
		style = xlwt.XFStyle()

		# font
		font = xlwt.Font()
		font.height = 300
		font.bold = True
		style.font = font
		worksheet.write(0, 0, title, style=style)
		worksheet.row(0).height_mismatch = True
		worksheet.row(0).height = 400

		for j in range(len(quarterly_data[i]["Registration"])):
			write_2d_list(worksheet, start_row, start_col, quarterly_data[i]["Registration"][j], j)
			start_row = start_row + len(quarterly_data[i]["Registration"][j])
		start_row = start_row + 3
		# print("sheet = ", sheet_list[i])
		# print(quarterly_data[i]["Registration"][0][0])
		# quarterly_data[i]["Purchase"][0].insert(0, quarterly_data[i]["Registration"][0][0])
		for j in range(len(quarterly_data[i]["Purchase"])):
			write_2d_list(worksheet, start_row, start_col, quarterly_data[i]["Purchase"][j], j)
			start_row = start_row + len(quarterly_data[i]["Purchase"][j])
		start_row = start_row + 3		
		for j in range(len(quarterly_data[i]["Purchase-new"])):
				write_2d_list(worksheet, start_row, start_col, quarterly_data[i]["Purchase-new"][j], j)
				start_row = start_row + len(quarterly_data[i]["Purchase-new"][j])
		start_row = start_row + 3		
		for j in range(len(quarterly_data[i]["Purchase-repeat"])):
				write_2d_list(worksheet, start_row, start_col, quarterly_data[i]["Purchase-repeat"][j], j)
				start_row = start_row + len(quarterly_data[i]["Purchase-repeat"][j])	
		start_row = start_row + 3

	wb.save("./Q1_2020.xls")                            
		

main()
