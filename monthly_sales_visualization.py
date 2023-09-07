import pandas as pd
import matplotlib.pyplot as plt

# read the Excel file into a pandas DataFrame
file_path = "/Agave_Spirits_Sales/Seznam faktur vystavených 04.09.23 cleaned.xlsx"
df = pd.read_excel(file_path)

# convert the "Dat.vystavení" column to datetime format
# handle variations in date formats by using the errors='coerce' parameter when converting the 'Dat.vystaveni' column to datetime.
# this parameter will handle errors by setting the problematic dates to NaT (Not a Timestamp)
df['Dat.vystavení'] = pd.to_datetime(df['Dat.vystavení'], format='%d.%m.%Y', errors='coerce')

###
##
# filter the data for the desired months (March to September)
df_filtered = df[(df['Dat.vystavení'].dt.month >= 3) & (df['Dat.vystavení'].dt.month <= 9)]
df_filtered['Month'] = df_filtered['Dat.vystavení'].dt.month

# group the data by month and sum the "Celkem s DPH" column
monthly_sales = df_filtered.groupby('Month')['Celkem s DPH'].sum()

# create a bar graph
plt.figure(figsize=(10, 6))

# customize x-axis labels to display month names
month_names = ['Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep']
ax = monthly_sales.plot(kind='bar', color='skyblue')

# set labels and title
plt.xlabel('Months')
plt.ylabel('Sales')
plt.title('Agave Spirits Mezcal Sales (March to September)')
plt.xticks(monthly_sales.index - 3, month_names)  # Subtract 3 to start from March

# annotate bars with their respective values
for i, v in enumerate(monthly_sales):
    ax.text(i, v, str(v), ha='center', va='bottom', fontsize=12)

# display graph
plt.tight_layout()
plt.show()
