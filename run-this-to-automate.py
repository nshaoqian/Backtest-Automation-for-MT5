#####################################  setup  #####################################
import datetime
import os
import xml.etree.ElementTree as ET
from bs4 import BeautifulSoup
import numpy as np
import pandas as pd

# directory
terminal_folder = r"D:\\Desktop\\TDX\\terminal64.exe"
report_folder = r"C:\Users\USER\AppData\Roaming\MetaQuotes\Terminal\BFF0607C79B1810AA4EFA95DE50F2A17\reports" # MT5 data folder
opt_ini_folder = r"D:\Desktop\TDX\Backtest Automation\optimization"
for_ini_folder = r"D:\Desktop\TDX\Backtest Automation\forecast"

def first_sunday(year):
    jan_1 = datetime.date(year, 1, 1)   
    weekday = jan_1.weekday()            # determine the day of the week of January 1st
    days_to_sunday = (6 - weekday) % 7   # find the number of days to the following Sunday
    first_sunday_date = jan_1 + datetime.timedelta(days=days_to_sunday) 
    
    return first_sunday_date

symbol = input("Enter symbol (eg XAUUSD): ")
year = int(input("Enter year: "))
period = int(input("Enter number of training days: "))

startdate = first_sunday(year)                                  # first Sunday of the year
enddate = first_sunday(year+1) + datetime.timedelta(days=-1)    # last Saturday before the next year
if year == 2023:
    enddate = datetime.date(2023, 7, 1)

numday = abs(startdate-enddate).days
numweek = numday //7 + 1

# documentation
week = []
forecast_start = []
forecast_end = []
training_start = []
training_end = []
tppointlist = []
slpointlist = []
lot = np.full((1, numweek), 0.1)
buyselllist = []
netprofitlist = []
capital = [1000]

######################### 1) write bat files for automations #########################

# write bat file for optimization automation
output_filename = "automate-optimization.bat"

with open(output_filename, "w") as output_file:
    for i in range(1, numweek + 1):
        ini_file_path = f"{opt_ini_folder}\\opt-week{i}.ini"
        line = f'"D:\\Desktop\\TDX\\terminal64.exe" /config:"{ini_file_path}"\n'
        output_file.write(line)

# write bat file for forecast automation
output_filename = "automate-forecast.bat"

with open(output_filename, "w") as output_file:
    for i in range(1, numweek + 1):
        ini_file_path = f"{for_ini_folder}\\for-week{i}.ini"
        line = f'"{terminal_folder}" /config:"{ini_file_path}"\n'
        output_file.write(line)

########### 2) write optimization ini files with different training periods ###########

start = startdate + datetime.timedelta(days=-period)
end = start + datetime.timedelta(days=period-1)  

# read template file
opt_template_path = "optimization/template-opt.txt"
with open(opt_template_path, "r") as opt_template:
    opt_template = opt_template.read()

for i in range(1,numweek+1):
    # write .ini files for optimization
    content = opt_template.replace("Symbol=", f"Symbol={symbol}") \
                          .replace("FromDate=", f"FromDate={start}") \
                          .replace("ToDate=", f"ToDate={end}") \
                          .replace("reports\opti-week", f"reports\opti-week{i}")

    output_file_path = os.path.join(opt_ini_folder, "opt-week"+str(i)+".ini")
    with open(output_file_path, "w") as output_file:
        output_file.write(content)

    week.append(i)
    training_start.append(start)
    training_end.append(end)

    start = start + datetime.timedelta(days=7)
    end = end + datetime.timedelta(days=7)    

###################### 3) start automation for optimization ##########################

os.system("automate-optimization.bat")
          
########## 4) write forecasting ini files with different forecasting periods ##########

start = startdate
end = start + datetime.timedelta(days=6) 

# read template file
for_template_path = "forecast/template-for.txt"
with open(for_template_path, "r") as for_template:
    for_template = for_template.read()

for i in range(1,numweek+1):
    # write .ini files for forecasting
    report_name = "opti-week"+str(i)+".xml"
    report_path = os.path.join(report_folder, report_name)
    tree = ET.parse(report_path)
    root = tree.getroot()
    
    TpPoint = root[2][0][1][10][0].text
    SlPoint = root[2][0][1][11][0].text
    BuyTrue_SellFalse = root[2][0][1][12][0].text
    
    content = for_template.replace("Symbol=", f"Symbol={symbol}") \
                          .replace("FromDate=", f"FromDate={start}") \
                          .replace("ToDate=", f"ToDate={end}") \
                          .replace("forecast-week", f"forecast-week{i}") \
                          .replace("TpPoints=", f"TpPoints={TpPoint}") \
                          .replace("SlPoints=", f"SlPoints={SlPoint}") \
                          .replace("BuyTrue_SellFalse=", f"BuyTrue_SellFalse={BuyTrue_SellFalse}")

    output_file_path = os.path.join(for_ini_folder, "for-week"+str(i)+".ini")
    with open(output_file_path, "w") as output_file:
        output_file.write(content)

    forecast_start.append(start)
    forecast_end.append(end)
    tppointlist.append(TpPoint)
    slpointlist.append(SlPoint)
    buyselllist.append("Buy" if BuyTrue_SellFalse == "true" else "Sell")

    start = start + datetime.timedelta(days=7)
    end = end + datetime.timedelta(days=7)    

###################### 5) start automation for forecasting ######################

os.system("automate-forecast.bat")

################# 6) extract weekly net profit from the htm files #################

# Parse the htm file in each week and extract the net profit
netProfit = []

for i in range(1, numweek + 1):
    report_name = "forecast-week"+str(i)+".htm"
    htm_file_path = os.path.join(report_folder, report_name)

    # Read the HTML content from the file
    with open(htm_file_path, "r", encoding="utf-16LE") as htm_file:
        html_content = htm_file.read()

    # Parse the HTML content using BeautifulSoup
    soup = BeautifulSoup(html_content, "html.parser")

    # Find the <td> element that contains the text "Total Net Profit:"
    total_net_profit_element = soup.find("td", string="Total Net Profit:")

    # Get the <td> element that contains the value you want
    net_profit_value_element = total_net_profit_element.find_next("td")

    # Extract the text from the element
    net_profit_value = net_profit_value_element.get_text()

    netprofitlist.append(float(net_profit_value.replace(" ", ""))/100)
    capital.append(capital[i-1]*(1+netprofitlist[i-1]/100))

capital.remove(capital[0])

def delete_files_in_folder(folder_path):
    try:
        # List all files in the folder
        file_list = os.listdir(folder_path)
        
        # Loop through the file list and delete each file
        for file_name in file_list:
            file_path = os.path.join(folder_path, file_name)
            if os.path.isfile(file_path):
                os.remove(file_path)
                print(f"Deleted: {file_path}")
                
        print("All files deleted successfully.")
    except Exception as e:
        print(f"An error occurred: {e}")

deletedecision = input("Delete reports? (y/n): ")
# deletedecision = "n" # uncomment this line to skip deleting reports
if deletedecision == "y":
    delete_files_in_folder(report_folder)

######################### 7) documentation #########################

header = ["week", "forecast start", "forecast end", "training start", "training end",
    "TP", "SL", "Lot", "Week Position Bias", "Profit(%)", "Starting Capital USD 1000"]

combined_array = np.vstack((week, forecast_start, forecast_end, training_start, training_end, tppointlist, slpointlist, lot, buyselllist, netprofitlist, capital))

df = pd.DataFrame(np.transpose(combined_array), columns = header)

folder_name = "results"
if not os.path.exists(folder_name):
    os.makedirs(folder_name)

# Construct the file path within the "results" folder
file_path = os.path.join(folder_name, f"{symbol}_{year}_{period}.xlsx")

# Create an Excel writer object
excel_writer = pd.ExcelWriter(file_path, engine="xlsxwriter")

# Write the DataFrame to the Excel file
df.to_excel(excel_writer, sheet_name="Sheet1", index=False)

# Get the xlsxwriter workbook and worksheet objects
workbook = excel_writer.book
worksheet = excel_writer.sheets["Sheet1"]

# Add a header format
header_format = workbook.add_format({
    "bold": True,
    "text_wrap": True,
    "valign": "top",
    "border": 1
})

# Set the column width and apply the header format
for col_num, value in enumerate(df.columns.values):
    worksheet.write(0, col_num, value, header_format)
    column_len = max(df[value].astype(str).apply(len).max(), len(value))
    worksheet.set_column(col_num, col_num, column_len + 2)

# Save the Excel file
excel_writer.save()

print(f"{file_path} generated successfully.")