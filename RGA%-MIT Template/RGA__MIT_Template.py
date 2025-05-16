import pandas as pd #used to load, process, and manipulate the data from the Excel file
import matplotlib.pyplot as plt #used to generate the plots and graphs in your code
import os #not directly used but, checks if the directory exists before saving a file or handles file paths in a platform-independent way
import matplotlib.dates as mdates #used to format the x-axis of your plot to display time-based data
from datetime import datetime # used to generate a timestamp for the filename when you save the plot
from openpyxl import load_workbook
import matplotlib.ticker as ticker

# 🔧 ============================== #
#     REQUIRED UPDATES FOR NEW FILE

excel_file = r'\\TRUENAS\PlasmaFlow\Staff\Emily H\PYTHON\RGA%-MIT Template\RGA%-MIT Template\RGA%-MIT Template\5-14-25 MIT TESTING-tab2.xlsx'
base_name = "5-14-25 MIT TESTING-tab2" # Base name used for versioned filenames and plot title
output_dir = "."  # Or your desired output path
data_ext = ".xlsx"
column_headers = ['Time','N2%', 'CH4%', 'H2%', 'O2%', 'CO%', 'CO2%', 'H2O%', 'C3H8%', 'N2corr CH4 Conv%']
plot_colors = ['tab:blue', 'tab:orange', 'tab:green', 'tab:pink', 'tab:red', 'tab:purple', 'tab:brown','tab:olive', 'tab:cyan', 'tab:gray']
y_axis1_limit = (0, 15)   # Left Y-axis 
y_axis2_limit = (30, 100)

# Set up the plot (this is where we define ax1)
fig, ax1 = plt.subplots(figsize=(12, 6))
plot_title = base_name
xlabel= 'Time'
ax2 = ax1.twinx()
ylabel_left= 'Species %'
ylabel_right= 'N2 Conv% & N2%'
start_time = pd.to_datetime('9:24:00 AM').time() #Define your start times
end_time = pd.to_datetime('12:40:00 PM').time() #Define your end times
# 🔧 ============================== #

# Define base name and output path
output_dir = "."  # Or your desired output path
data_ext = ".xlsx"

# 🔁 Find next available version number
version = 1
while True:
    versioned_filename = os.path.join(output_dir, f"{base_name}-V{version}{data_ext}")
    if not os.path.exists(versioned_filename):
        break 
    version += 1

# Load the workbook
wb = load_workbook(excel_file)

df = pd.read_excel(excel_file, engine='openpyxl', skiprows=39)#use subset = df.iloc[0:100] (first 100 rows after skipping the first 36 in Excel)
print("Columns in the DataFrame:", df.columns.tolist())
# Load the Excel file, skipping the first 36 rows, and let pandas use the 37th row as headers
df['Time'] = pd.to_datetime(df['Time']) #Convert the 'Time' column to datetime (if not already)
df = df[df['Time'].dt.time.between(start_time, end_time)] # 🔥 Filter rows within your time range

# Time formatting for X-axis
time_format = mdates.DateFormatter('%I:%M:%S %p')  # Display in Hour:Minute:Second & 12 hr format
ax1.xaxis.set_major_formatter(time_format)

# Loop through columns and colors
for col_name, color in zip(column_headers, plot_colors): #this creates the loop to pair column headers with the plot colors
    if col_name == 'Time': #says to skip this iteration if x = " " 
        continue
    elif col_name == 'N2corr CH4 Conv%': #if the columns name is x then skip plotting (else if)
        ax2.plot(df['Time'], df[col_name], label=col_name, linewidth=2, color=color)
    else: #says to plot all other data on a different axis
        ax1.plot(df['Time'], df[col_name], label=col_name, linewidth=2, color=color)

# Set x-axis intervals
ax1.xaxis.set_major_locator(mdates.MinuteLocator(interval=2))  # Change interval to whatever works
ax1.xaxis.set_minor_locator(mdates.MinuteLocator(interval=1))  # Optional: Add minor ticks every minute

# Rotate the x-axis labels for better visibility, could add #fontweight='bold'if needed
plt.setp(ax1.get_xticklabels(), rotation=45, ha='right', fontsize='8') 

# Set y-axis intervals
ax1.yaxis.set_major_locator(ticker.MultipleLocator(2))  # Change interval to whatever works
ax2.yaxis.set_major_locator(ticker.MultipleLocator(10))  # Change interval to whatever works

# Title
title_kwargs = {
    'fontsize': 14,
    'fontweight': 'bold',
    'color': 'black',
    'loc': 'center',
    'pad': 15,
    'fontname': 'Arial'
}

label_kwargs = {
    'fontsize': 11,
    'fontweight': 'bold',
    'fontname': 'Arial'
}

# Combine legends from both y-axes
lines1, labels1 = ax1.get_legend_handles_labels()
lines2, labels2 = ax2.get_legend_handles_labels()
all_lines = lines1 + lines2
all_labels = labels1 + labels2

ax1.set_ylim(*y_axis1_limit)
ax2.set_ylim(*y_axis2_limit)

# Place the combined legend outside the plot
legend = ax1.legend(
    all_lines, all_labels,           # Combine line objects and their labels from ax1 and ax2
    loc='upper center',              # Place legend above the plot, centered
    bbox_to_anchor=(0.5, -0.2),      # Fine-tune its position (centered, below plot)
    ncol=7,                          # Spread legend items across 7 columns
    fontsize=10                      # Set font size of legend text
)

# Legend text
for text in legend.get_texts():
    text.set_fontweight('bold')
    text.set_fontname('Arial')

ax1.grid(True, which='both', linestyle='--', linewidth=0.5, alpha=0.7)
plt.title(plot_title, **title_kwargs)
plt.tight_layout()
# Style tick labels (x-axis and both y-axes)
plt.setp(ax1.get_xticklabels(), fontsize=6, fontweight='bold', fontname='Arial')
plt.setp(ax1.get_yticklabels(), fontsize=8, fontweight='bold', fontname='Arial')
plt.setp(ax2.get_yticklabels(), fontsize=8, fontweight='bold', fontname='Arial')


# ✅ Save Excel data to the already-versioned file
df.to_excel(versioned_filename, index=False)
print(f"✅ Data saved as '{versioned_filename}'")

# Define base name and output path
output_dir = "."  # Or your desired output path
plt_ext = ".png"

# 🔁 Find next available version number
version = 1
while True:
    versioned_plot_filename = os.path.join(output_dir, f"{base_name}-V{version}{plt_ext}")
    if not os.path.exists(versioned_plot_filename):
        break
    version += 1

# ✅ Save plot with same version number
plt.savefig(versioned_plot_filename)
print(f"✅ Plot saved as '{versioned_plot_filename}'")

plt.show()

