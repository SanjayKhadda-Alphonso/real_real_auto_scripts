import csv
import pandas as pd
import datetime
import network as net

file_path = "/Users/sanjaykhadda/Desktop/real_real_scripts/reduce_impressions/0309_0315_v2.csv"


df_spend = pd.read_csv(file_path , engine='python', encoding='UTF-8', index_col=False)
# df_spend = df_spend.drop(df_spend.columns[0],axis=1)

df_spend["imp_over_spend_reduced"] = df_spend["Impression Share"]/df_spend["Spend Share"]
df_spend["imp_reduced"] = df_spend["Impressions"]
natloc_factor = 1.5
national_factor = 2.2
df_spend["imp_share_reduced"] = df_spend["Impression Share"]
done = False
total_impressions = df_spend["Impressions"].sum()
print("total_impressions = ", total_impressions)
while done == False:
	done = True
	for i in range(len(df_spend["network"])):
		if df_spend.Spend[i] != 0:
			if df_spend.network[i] in net.natloc_or_national.keys():
				if net.natloc_or_national[df_spend.network[i]] == "Natloc":
					#print("XX ", df_spend.network[i])
					if round(df_spend.imp_over_spend_reduced[i], 2) > natloc_factor:
						#print("1 ", df_spend.network[i])
						df_spend.imp_reduced[i] = df_spend.imp_reduced[i]*natloc_factor/df_spend.imp_over_spend_reduced[i]
						done = False
				elif net.natloc_or_national[df_spend.network[i]] == "National":
					if round(df_spend.imp_over_spend_reduced[i],2) > national_factor:
						#print("2 ", df_spend.network[i])
						df_spend.imp_reduced[i] = df_spend.imp_reduced[i]*national_factor/df_spend.imp_over_spend_reduced[i]
						done = False
			else:
				# print("Network ",  df_spend.network[i], " not found in the mapping file. Assuming it's natloc")
				if round(df_spend.imp_over_spend_reduced[i],2) > natloc_factor:
					#print("3 ", df_spend.network[i])
					df_spend.imp_reduced[i] = df_spend.imp_reduced[i]*natloc_factor/df_spend.imp_over_spend_reduced[i]
					done = False
	print(df_spend.imp_over_spend_reduced)
	# print(df_spend.imp_over_spend_reduced)
	total_impressions = df_spend["imp_reduced"].sum()
	df_spend["imp_share_reduced"] = df_spend["imp_reduced"]*100/total_impressions
	df_spend["imp_over_spend_reduced"] = df_spend["imp_share_reduced"]/df_spend["Spend Share"]
df_spend["CPM_Reduced"] = 1000*(df_spend["Spend"]/(5*df_spend["imp_reduced"]))

# For copying to put in the reducing script
copy = "("
for i in range(len(df_spend["network"])):
	if df_spend.imp_reduced[i]/df_spend.Impressions[i] != 1 and df_spend.imp_reduced[i] > 0:
		#print("Factor for ", df_spend.network[i], " is ", df_spend.imp_reduced[i]/df_spend.impressions[i])
		copy += "(\""+ df_spend.network[i] + "\", "+str(df_spend.imp_reduced[i]/df_spend.Impressions[i])+"),"
copy = copy[:-1]
copy += ")"

print(copy)

df_spend.to_csv("impression_reduced.csv", index=False)
