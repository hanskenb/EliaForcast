import pandas as pd
import matplotlib.pyplot as plt
import requests
from datetime import datetime, timedelta
from io import BytesIO

def fetch_excel(url):
    response = requests.get(url)
    data = pd.read_excel(BytesIO(response.content))
    # Find the row index where 'DateTime' is located
    datetime_row_idx = data.index[data.apply(lambda row: row.str.contains('DateTime', na=False).any(), axis=1)][0] + 1
    return pd.read_excel(BytesIO(response.content), skiprows=datetime_row_idx)

def extract_forecast(data):
    return data[['DateTime', 'Most recent forecast [MW]']].dropna()  # Dropping rows with NaN values

def create_chart():
    # Generating URLs
    from_date = datetime.now().strftime('%Y-%m-%d')
    to_date = (datetime.now() + timedelta(days=7)).strftime('%Y-%m-%d')

    wind_url = f"https://griddata.elia.be/eliabecontrols.prod/interface/fdn/download/windweekly/currentselection?dtFrom={from_date}&dtTo={to_date}&sourceID=1&forecast=wind"
    solar_url = f"https://griddata.elia.be/eliabecontrols.prod/interface/fdn/download/solarweekly/currentselection?dtFrom={from_date}&dtTo={to_date}&sourceID=1&forecast=solar"

    # Fetch and process data
    wind_data = fetch_excel(wind_url)
    solar_data = fetch_excel(solar_url)

    wind_forecast = extract_forecast(wind_data)
    solar_forecast = extract_forecast(solar_data)

    # Combining data for chart
    combined_forecast = pd.merge(wind_forecast, solar_forecast, on='DateTime', suffixes=('_wind', '_solar'))
    combined_forecast['Cumulative Forecast'] = combined_forecast['Most recent forecast [MW]_wind'] + combined_forecast['Most recent forecast [MW]_solar']

    # Plotting
    plt.figure(figsize=(10, 6))
    plt.plot(combined_forecast['DateTime'], combined_forecast['Most recent forecast [MW]_wind'], label='Wind Forecast', color='blue')
    plt.plot(combined_forecast['DateTime'], combined_forecast['Most recent forecast [MW]_solar'], label='Solar Forecast', color='orange')
    plt.plot(combined_forecast['DateTime'], combined_forecast['Cumulative Forecast'], label='Cumulative Forecast', color='green')

    plt.xlabel('DateTime')
    plt.ylabel('Forecast (MW)')
    plt.title('Wind and Solar Energy Forecast')
    plt.legend()
    plt.show()

# Call the function without arguments
create_chart()
