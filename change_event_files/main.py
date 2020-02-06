import pandas as pd
import os.path

def modify_file(path, file, file_directory):
    df = pd.read_csv(path, engine='python', delim_whitespace=False, encoding='UTF-8')
    for i in range(len(df.user_id)):
    	if df.event_name[i] == "Purchase":
            df.event_name[i] = df.event_name[i] + "-" + df.user_type[i]
    df.to_csv(file_directory + "changed_files/" + file, index=False)


file_directory = "./"
save_directory = ""
for file in os.listdir(file_directory):
    if file.endswith(".csv"):
        path = os.path.join(file_directory, file)
        modify_file(path, file, file_directory)
