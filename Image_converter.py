import pprint as pp
import pandas as pd
from urllib.request import urlretrieve
from openpyxl import load_workbook
from openpyxl.worksheet.properties import WorksheetProperties as wp

#  List of image srcs
Primary_list = [
'img src list here'
]

#  List of skus
mfg_list = [
'sku list here'
 ]

# Create a dictionary from the two lists for a loop
img_dict = {mfg_list[i]: Primary_list[i] for i in range(len(Primary_list))}
pp.pprint(img_dict)

# Creating a folder variable for output
output_directory = r'path\to\desktop\folder'

# Looping through the dictionary and creating .jpgs from the urls and loading the file names into a list
filename_list = []
for mfg, url in img_dict.items():
    file_name = mfg + '_Primary.jpg'
    urlretrieve(url, output_directory + f"\{file_name}")
    filename_list.append(file_name)

# Creating a dataframe from the file name list
file_df = pd.DataFrame({'Image_File_Name': filename_list})
pp.pprint(file_df)

#  Writing the dataframe to an excel worksheet
path = 'excel_file.xlsx'
excel_wb = load_workbook(path)
with pd.ExcelWriter(path) as writer:
    writer.book = excel_wb
    file_df.to_excel(writer, sheet_name='File_Data', index=False)
    file_sheet = writer.sheets['File_Data']
    file_sheet.sheet_properties.tabColor = 'FFFF00'

print('Run Complete!')