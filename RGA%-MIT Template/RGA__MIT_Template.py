import pandas as pd #used to load, process, and manipulate the data from the Excel file
import matplotlib.pyplot as plt #used to generate the plots and graphs in your code
import os #not directly used but, checks if the directory exists before saving a file or handles file paths in a platform-independent way
import matplotlib.dates as mdates #used to format the x-axis of your plot to display time-based data
from datetime import datetime # used to generate a timestamp for the filename when you save the plot
from openpyxl import load_workbook
import matplotlib.ticker as ticker

# 🔧 ============================== #
#     REQUIRED UPDATES FOR NEW FILE

excel_file = r'\\TRUENAS\PlasmaFlow\Staff\Emily H\PYTHON\RGA% Code Template\RGA% Code Template\4-14-25 TC2-R2 TESTING-%.xlsx'
base_name = "4-14-25 TC2-R2 TESTING-%" # Base name used for versioned filenames and plot title
output_dir = "."  # Or your desired output path
data_ext = ".xlsx"
column_headers = ['TIME', 'H2%', 'CH4%', 'C2H2%', 'C2H4%', 'C3H8%', 'C3H6%']
plot_colors = ['tab:blue', 'tab:orange', 'tab:green', 'tab:pink', 'tab:red', 'tab:purple']
y_axis1_limit = (0, 200)   # Left Y-axis for TC2–TC6
y_axis2_limit = (0, 350)

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

# Load the Excel file, skipping the first 36 rows, and let pandas use the 37th row as headers
df = pd.read_excel(excel_file, engine='openpyxl', skiprows=36)  #use subset = df.iloc[0:100] (first 100 rows after skipping the first 36 in Excel)
df['Time'] = pd.to_datetime(df['Time']) #Convert the 'Time' column to datetime (if not already)
start_time = pd.to_datetime('3:55:00 PM').time() #Define your start and end times
end_time = pd.to_datetime('4:32:00 PM').time() #Define your start and end times
df = df[df['Time'].dt.time.between(start_time, end_time)] # 🔥 Filter rows within your time range

# Set up the plot (this is where we define ax1)
fig, ax1 = plt.subplots(figsize=(12, 6))

# Time formatting for X-axis
time_format = mdates.DateFormatter('%I:%M:%S %p')  # Display in Hour:Minute:Second & 12 hr format
ax1.xaxis.set_major_formatter(time_format)

# Plot each of the columns S to Y (index 18 to 24)
ax2 = ax1.twinx()
ax2.plot(df['Time'], df['H2%'], label='H2%', linewidth=2, color='tab:blue')
ax1.plot(df['Time'], df['CH4%'], label='CH4%', linewidth=2, color='tab:red')
ax1.plot(df['Time'], df['C2H2%'], label='C2H2%', linewidth=2, color='tab:pink')
ax1.plot(df['Time'], df['C2H4%'], label='C2H4%', linewidth=2, color='tab:cyan')
ax1.plot(df['Time'], df['C3H8%'], label='C3H8%', linewidth=2, color='tab:gray')
ax1.plot(df['Time'], df['C2H6%'], label='C2H6%', linewidth=2, color='tab:green')

# Set x-axis intervals
ax1.xaxis.set_major_locator(mdates.MinuteLocator(interval=2))  # Change interval to whatever works
ax1.xaxis.set_minor_locator(mdates.MinuteLocator(interval=1))  # Optional: Add minor ticks every minute

# Rotate the x-axis labels for better visibility, could add #fontweight='bold'if needed
plt.setp(ax1.get_xticklabels(), rotation=45, ha='right', fontsize='8') 

# Manually setting Y-axis ranges
ax1.set_ylim(0, 20) 
ax2.set_ylim(0, 100)

# Set y-axis intervals
ax1.yaxis.set_major_locator(ticker.MultipleLocator(2))  # Change interval to whatever works
ax2.yaxis.set_major_locator(ticker.MultipleLocator(10))  # Change interval to whatever works


# Labels 
# Title
plt.title('4-14-25 R1H PRODUCTION - RGA %', fontsize=14, fontweight='bold', color='black', loc='center', pad=15, fontname='Arial') #where pad = the pixel spacing between the graph and title
#X axis
ax1.set_xlabel('Time', fontsize=10, fontweight='bold', fontname='Arial')
#Y1 axis
ax1.set_ylabel('Species %', fontsize=11, fontweight='bold', fontname='Arial')
#Y2 axis
ax2.set_ylabel('H2 %', fontsize=11, fontweight='bold', fontname='Arial')

# Combine legends from both y-axes
lines1, labels1 = ax1.get_legend_handles_labels()
lines2, labels2 = ax2.get_legend_handles_labels()
all_lines = lines1 + lines2
all_labels = labels1 + labels2

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

