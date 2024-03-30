import requests
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
import pandas as pd
import random
import time

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/97.0.4692.99 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/97.0.4692.99 Safari/537.36",
]

# Function to scrape weather data for a given date and station code
def scrape_weather_data_for_date(url):
    retries = 3  # Number of retries
    backoff_factor = 2  # Backoff factor for exponential backoff
    for attempt in range(retries):
        try:
            # Rotate user agent
            headers = {'User-Agent': random.choice(USER_AGENTS)}
            
            # Send HTTP request
            response = requests.get(url, headers=headers, timeout=10)
            response.raise_for_status()  # Raise an exception for 4xx and 5xx status codes
            
            # Parse HTML
            soup = BeautifulSoup(response.content, 'html.parser')
            
            # Extract station name
            station_name_elem = soup.select_one('#inner-content > div.region-content-top > app-dashboard-header > div.dashboard__header.small-12.ng-star-inserted > div > div.station-nav > a')
            station_name = station_name_elem.text.strip() if station_name_elem else 'N/A'
            
            # Extract elevation
            elevation_elem = soup.select_one('#inner-content > div.region-content-top > app-dashboard-header > div.dashboard__header.small-12.ng-star-inserted > div > div.sub-heading > span > strong:nth-child(1)')
            elevation = elevation_elem.text.strip() if elevation_elem else 'N/A'
            
            # Extract weather data for each day
            weather_data = []
            for row in soup.select('#main-page-content > div > div > div > lib-history > div.history-tabs > lib-history-table > div > div > div > table > tbody > tr'):
                try:
                    day = row.select_one('td:nth-child(1)').text.strip()  # Extract day
                    temp_max = row.select_one('td:nth-child(2) lib-display-unit > span > span.wu-value.wu-value-to').text.strip()
                    temp_avg = row.select_one('td:nth-child(3) lib-display-unit > span > span.wu-value.wu-value-to').text.strip()
                    temp_min = row.select_one('td:nth-child(4) lib-display-unit > span > span.wu-value.wu-value-to').text.strip()
                    precip_total_elem = row.select_one('td:nth-child(16) lib-display-unit > span > span.wu-value.wu-value-to')
                    precip_total = precip_total_elem.text.strip() if precip_total_elem else 'N/A'
                    
                    data_points = [
                        station_name,
                        elevation,
                        day,
                        temp_max,
                        temp_avg,
                        temp_min,
                        precip_total
                    ]
                except AttributeError:
                    # If any of the required data points are not found, append "N/A" to the data list
                    data_points = ['N/A'] * 7
                
                weather_data.append({'Date': day, 'Data': data_points})
            
            return station_name, elevation, weather_data
        except (requests.exceptions.RequestException, ValueError) as e:
            print(f"Error fetching page: {e}")
            if attempt < retries - 1:
                # Implement exponential backoff
                delay = backoff_factor ** attempt
                print(f"Retrying in {delay} seconds...")
                time.sleep(delay)  # Add a delay before retrying
            else:
                print("Exceeded maximum number of retries.")
                return None, None, []

# Function to generate date range by month in reverse order
def generate_monthly_date_range(start_date, end_date):
    current_date = end_date.replace(day=1)  # Start from the first day of the end month
    while current_date >= start_date:
        yield current_date
        # Move to the first day of the previous month
        if current_date.month == 1:
            current_date = current_date.replace(year=current_date.year - 1, month=12)
        else:
            current_date = current_date.replace(month=current_date.month - 1)

# Define start and end dates
start_date = datetime(2016, 1, 1)
end_date = datetime(2023, 12, 1)

# Define base URL
base_url = 'https://www.wunderground.com/dashboard/pws/'

# List of station codes
station_codes = [
    'IVALTO2',
    'IGARMISC34',
    'IMITTE82',
    'IBALDE3',
    'IOBERA47',
    'IBAYERNR33',
    'IOBERA42',
    'IBERCH21',
    'ILOBAR3',
    'IGANGN11',
    'IRUERA5',
    'ILECH44',
    'ITABRUNI5',
    'ILOURT1',
    'IROURE1',
    'ILESAL6',
    'IVALLEDA13',
    'IMORGE4'
]

# Initialize empty lists to store data
data = []

# Loop through each station code
for station_code in station_codes:
    # Loop through each month in the date range in reverse order
    for date in generate_monthly_date_range(start_date, end_date):
        # Construct URL for the first day of the current month and station code
        url = f'{base_url}{station_code}/table/{date.strftime("%Y-%m-%d")}/{date.strftime("%Y-%m-%d")}/monthly'
        
        # Scrape weather data for the current date and station code
        station_name, elevation, weather_data = scrape_weather_data_for_date(url)
        
        # Extract relevant data
        for entry in weather_data:
            data.append({
                'Station': station_code,
                'Location': entry['Data'][0],
                'Elevation': entry['Data'][1],
                'Date': entry['Date'],
                'Temp Max': entry['Data'][3],
                'Temp Avg': entry['Data'][4],
                'Temp Min': entry['Data'][5],
                'Precip Total': entry['Data'][6]
            })
# Create a DataFrame from the data
df = pd.DataFrame(data)
# Write DataFrame to Excel file
output_file = 'weather_data.xlsx'
df.to_excel(output_file, index=False)