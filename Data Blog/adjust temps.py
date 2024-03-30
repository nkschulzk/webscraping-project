import openpyxl
import pandas as pd

# Load data from Excel file
excel_file = 'weather_data.xlsx'
workbook = openpyxl.load_workbook(excel_file)
worksheet = workbook.active
rows = worksheet.iter_rows(min_row=2, values_only=True)

# Define column names
column_names = ['Station', 'Location', 'Elevation', 'Date', 'Temp Max', 'Temp Avg', 'Temp Min', 'Precip Total']

# Create DataFrame with specified column names
df = pd.DataFrame(rows, columns=column_names)

# Create DataFrame for station elevation data
resort_data = pd.DataFrame({
    'ResortName': ['Garmisch-Partenkirchen', 'Oberstdorf', 'Balderschwang', 'Oberammergau', 
                 'Reit im Winkl', 'Oberaudorf', 'Berchtesgaden', 'Valle Nevado', 
                 'Alpensia (South Korea)', 'Andermatt-Sedrun Sport AG', 'Ski Arlberg', 
                 'Skirama Dolomiti Adamello Brenta', 'Sestriere', 'Les 3 Vallées', 'Verbier4Vallées', 'Cervinia'],
    'Station': ['IGARMISC34', 'IMITTE82', 'IBALDE3', 'IOBERA47', 'IBAYERNR33',
                'IOBERA42', 'IBERCH21', 'ILOBAR3', 'IGANGN11', 'IRUERA5',
                'ILECH44', 'ITABRUNI5', 'IROURE1', 'ILESAL6', 'IVALLEDA13', 'IVALTO2'],
    'resort_elev': [2303, 995, 1060, 834, 2651, 1637, 1998, 224, 20, 5062, 4869,
                    2732, 3271, 1270, 2392, 6722]
})

# Merge the dataframes
df = pd.merge(df, resort_data, on='Station', how='left')
df = df[df['ResortName'] != '']# Convert 'Elevation' column to numeric
df['Elevation'] = pd.to_numeric(df['Elevation'], errors='coerce')

# Convert feet to meters and calculate elevation change
df['Elevation'] *= 0.3048  # Conversion from feet to meters
df['elev_change'] = df['resort_elev'] - df['Elevation']

# Calculate temperature adjustment factor assuming a 5.4(F) degree drop for every thousand meters of elevation gain
df['temp_factor'] = -(df['elev_change'] / 1000) * 5.4

# Convert 'Temp Avg' and 'temp_factor' columns to numeric
df['Temp Avg'] = pd.to_numeric(df['Temp Avg'], errors='coerce')
df['temp_factor'] = pd.to_numeric(df['temp_factor'], errors='coerce')

# Calculate adjusted temperature
df['adj_temp'] = df['Temp Avg'] + df['temp_factor']

#Remove "Forecast for " from the location column
df['Location'] = df['Location'].str.replace('Forecast for ', '')

df['below_32'] = (df['adj_temp'] < 32).astype(int)
df['Precip Total'] = pd.to_numeric(df['Precip Total'], errors='coerce')
df['Date'] = pd.to_datetime(df['Date'], format='%m/%d/%Y')

#check for snow days
df['snowday'] = ((df['adj_temp'] < 32) & (df['Precip Total'] > 0)).astype(int)
#Drop the na rows in
monthly_avg_temp = df.groupby(['Station', df['Date'].dt.month])['adj_temp'].transform('mean')
# Fill NaN values in 'adj_temp' with monthly average temperature
df['adj_temp'].fillna(monthly_avg_temp, inplace=True)

#create a table with avg snowdays and avg days below freezing
df['Year'] = pd.to_datetime(df['Date']).dt.year
resort_year_grouped = df.groupby(['Location', 'Year'])
snow_days_per_year = resort_year_grouped['snowday'].sum().reset_index()
snow_days_per_year = snow_days_per_year.rename(columns={'snowday': 'Snow Days'})
print(snow_days_per_year)

output_file = 'adj_weather_data.xlsx'
df.to_excel(output_file, index=False)