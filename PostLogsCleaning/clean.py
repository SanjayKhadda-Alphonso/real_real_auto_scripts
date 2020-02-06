import csv
import dma
import pandas as pd
import datetime
import sys
import math


#file_path = "/Users/sanjaykhadda/Downloads/TheRealReal_Local_or.csv"
#file_path = "/Users/sanjaykhadda/Downloads/TheRealReal_National_930-106.csv"
#file_path = "/Users/sanjaykhadda/Downloads/national.csv"
#local_or_nat = "national"
#remove_LIFA = "y"
# file_path_national = "/Users/sanjaykhadda/Downloads/TheRealReal_National\ Media\ Log\ File\ 11.4.19\ -\ 11.10.19.csv "
# file_path_local = "/Users/sanjaykhadda/Downloads/TheRealReal_Local\ Media\ Log\ File\ 11.4.19\ -\ 11.10.19.csv"


file_path_national = input("Enter file path for national file: ")
file_path_local = input("Enter file path for local file: ")
remove_LIFA = input("Remove LIFA(y/n) : ")
re_run = False
while remove_LIFA not in ["Y", "y", "N", "n", "Yes", "No", "no", "yes"]:
	print("please enter valid response")
	remove_LIFA = input("Remove LIFA(y/n) : ")
if remove_LIFA in ["Y", "y", "Yes", "yes"]:
	remove_LIFA = True
else:
	remove_LIFA = False
# networks_to_remove_national = input("Enter Comma seperated networks to be removed from national eg. LIFA,ADS,OXY : ")

while "National" in file_path_local or "Local" in file_path_national:
	print("\n\nPlease check the paths, have you exchanged the local path with national or vice versa?")
	file_path_national = input("Enter file path for national file: ")
	file_path_local = input("Enter file path for local file: ")

file_path_national = file_path_national.replace("\\", "").strip()
file_path_local = file_path_local.replace("\\", "").strip()
file_name_national = "cleaned_national_file_" + file_path_national.split("/")[-1].replace(" ", "_")
file_name_local = "cleaned_local_file_" + file_path_local.split("/")[-1].replace(" ", "_")
with open(file_path_national) as f:
    print(f)
with open(file_path_national) as f:
	print(f)
df_national = pd.read_csv(file_path_national, engine='python', skipfooter=1, delim_whitespace=False, encoding='latin1')
df_local = pd.read_csv(file_path_local, engine='python', skipfooter=1, delim_whitespace=False, encoding='latin1')

pd.set_option('mode.chained_assignment', None)

#Local file processing
print("\n\n\n------------------------------PROCESSING LOCAL FILE------------------------------")
print('Changing column names...')
print(df_local.columns)
if 'CreativeID.1' in df_local.columns:
	df_local.rename(columns={'CreativeID':'Time',
	                      'CreativeID.1':'CreativeID'},
	             inplace=True)

df_local.rename(columns={'ï»¿Date': 'Date',
						'Time Zone':'TimeZone',
                      'Program Aired':'ProgramAired',
                      'Air Type':'RealRealAdType'},
             inplace=True)


print('Trimming Column names...')
for i in range(len(df_local.Date)):
	df_local.MarketName[i] = df_local.MarketName[i].strip()
	df_local.Time[i] = df_local.Time[i].strip()
	df_local.CreativeName[i] = df_local.CreativeName[i].strip()
	df_local.RealRealAdType[i] = df_local.RealRealAdType[i].strip()
	df_local.CreativeID[i] = df_local.CreativeID[i].strip()


print("Putting DMA Codes...")
print("\n\n\n|PUT DMA CODE FOR THESE MANUALLY(Also add the mapping in dma.py)|")
# print(df_local.MarketCode[0])
# print("len = ", len(df_local.MarketCode))
not_in_mapping = []
for i in range(len(df_local.MarketCode)):
	# print(df_local.MarketCode[i])
	# #print(df_local.MarketName[i])
	# print(math.isnan(df_local.MarketCode[i]))
	if df_local.MarketCode[i] != "" and math.isnan(df_local.MarketCode[i]) == False:
		# print("here")
		if str(int(df_local.MarketCode[i])) in dma.dma_code_to_code.keys():
			df_local.MarketCode[i]  = dma.dma_code_to_code[str(int(df_local.MarketCode[i]))]
		else:
			print("dma na = ", df_local.MarketCode[i])
	elif df_local.MarketName[i] in dma.dma_name_to_code.keys():
		df_local.MarketCode[i]  = str(int(dma.dma_name_to_code[df_local.MarketName[i]]))
	else:
		re_run = True
		not_in_mapping.append(df_local.MarketName[i])

print("|---------------------------------------------------------------|\n\n\n")

if re_run == True:
	print("Please re run the program after adding the mapping in dma.py")
	print(list(dict.fromkeys(not_in_mapping)), sep = "\n")
	sys.exit()

print("Trimming spaces in station name and removing TV Suffix and keeping only part before first '-' and first 4 letters...")
for i in range(len(df_local.Station)):
	df_local.Station[i] = df_local.Station[i].replace(" ", "")
	stationName = df_local.Station[i].split("-")
	df_local.Station[i] = stationName[0]
	df_local.Station[i] = df_local.Station[i][:4]
#print(df_local.Station)

print("Changing time format to HH:MM XM")
for i in range(len(df_local.Time)):
	if 'M' not in df_local.Time[i] and 'm' not in df_local.Time[i]:
		# print(df_local.Time[i], "iske liye")
		df_local.Time[i]= datetime.datetime.strptime(df_local.Time[i], '%H:%M:%S').strftime('%I:%M %p')
print("Writing csv file")
df_local.to_csv(file_name_local, index=False)
print("File created as " + file_name_local)
print("------------------------------PROCESSED LOCAL FILE------------------------------")

#national file processing
print("\n\n\n------------------------------PROCESSING NATIONAL FILE------------------------------")

print('Renaming columns...')
df_national.rename(columns={'ï»¿Date': 'Date',
					'Time Zone':'TimeZone',
                      'Cover %':'CoverPct',
                      'Program Aired':'ProgramAired',
                      'Air Type':'AirType',
                      'Local':'RealRealAdType',
                      },
             inplace=True)
print('Trimming Column names and rearranging the order...')
print(df_national.columns)
#print(df_national.Date[0])

df_national = df_national[['Date', 'Time', 'TimeZone', 'Network', 'RealRealAdType', 'ProgramAired', 'CreativeID', 'CreativeName', 'CreativeLength', 'AirType', 'Spend', 'CoverPct']]

for i in range(len(df_national.Date)):
	df_national.Date[i] = str(df_national.Date[i]).strip()
	df_national.Time[i] = str(df_national.Time[i]).strip()
	df_national.TimeZone[i] = str(df_national.TimeZone[i]).strip()
	df_national.Network[i] = df_national.Network[i].strip()
	df_national.ProgramAired[i] = df_national.ProgramAired[i].strip()
	df_national.CreativeID[i] = df_national.CreativeID[i].strip()
	df_national.CreativeName[i] = df_national.CreativeName[i].strip()
	df_national.CreativeLength[i] = str(df_national.CreativeLength[i]).strip()
	df_national.AirType[i] = df_national.AirType[i].strip()
	df_national.Spend[i] = str(df_national.Spend[i]).strip()
	df_national.CoverPct[i] = str(df_national.CoverPct[i]).strip()




print("Adding National and Natloc and removing L from local networks...")
for i in range(len(df_national.RealRealAdType)):
	if str(df_national.RealRealAdType[i]).strip() == "LOCAL":
		df_national.RealRealAdType[i] = "national_local"
		df_national.Network[i] = df_national.Network[i][:-1]
	else:#if str(df_national.RealRealAdType[i]).strip().isspace():
			df_national.RealRealAdType[i] = "National"

	# else:
	# 	print("Found ", str(df_national.RealRealAdType[i]).strip(), " at i = ", i, " cannot proceed further")
	# 	exit()


original = len(df_national.RealRealAdType)
print("Removing dual feed and satellite rows...")
for i in range(len(df_national.ProgramAired)):
	if df_national.ProgramAired[i] == "Dual Feed" or df_national.AirType[i] == "Satellite":
		df_national.drop(index = i, inplace = True)
		i = i - 1


print("Removed ", original - len(df_national.RealRealAdType), " rows. Earlier there were ", original, " rows now there are ",  len(df_national.RealRealAdType))

if remove_LIFA:
	print("Removing rows of LIFA")
	df_national = df_national[df_national.Network != "LIFA"]


print("Writing csv file...")
df_national.to_csv(file_name_national, index=False)
print("File created as ", file_name_national)
print("------------------------------PROCESSED NATIONAL FILE------------------------------\n\n\n")
