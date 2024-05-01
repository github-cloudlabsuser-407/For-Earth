import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import matplotlib.pyplot as plt


# Define start and end date
start_date = datetime(2023, 1, 1)
end_date = start_date + timedelta(days=180)

# Generate date range
date_range = pd.date_range(start_date, end_date, freq='15min')

# Generate random values for each column
voltage = np.random.uniform(100, 200, len(date_range))
current = np.random.uniform(130, 200, len(date_range))
frequency = np.random.normal(50, 0.1, len(date_range))
phase_angle = np.random.uniform(25, 35, len(date_range))
power_factor = np.random.uniform(0.85, 1, len(date_range))
temperature = np.random.uniform(35, 45, len(date_range))
humidity = np.random.uniform(0, 100, len(date_range))
wind_speed = np.random.uniform(0, 30, len(date_range)) # assuming max wind speed of 30 m/s
solar_radiation = np.random.uniform(0, 1000, len(date_range)) # assuming max solar radiation of 1000 W/m²
load_capacity = np.random.uniform(0, 100, len(date_range))
line_load = np.random.uniform(0, 100, len(date_range))
insulation_resistance = np.random.uniform(0, 100, len(date_range)) # assuming max insulation resistance of 100 MΩ
leakage_current = np.random.uniform(0, 10, len(date_range)) # assuming max leakage current of 10 mA

# Calculate power values
active_power = voltage * current * np.cos(np.deg2rad(phase_angle)) / 1e6 # in MW
reactive_power = voltage * current * np.sin(np.deg2rad(phase_angle)) / 1e6 # in MVAR
apparent_power = voltage * current / 1e6 # in MVA

# Create DataFrame
df = pd.DataFrame({
    'Timestamp': date_range,
    'Voltage (kV)': voltage,
    'Current (A)': current,
    'Frequency (Hz)': frequency,
    'Phase Angle (degrees)': phase_angle,
    'Power Factor': power_factor,
    'Active Power (MW)': active_power,
    'Reactive Power (MVAR)': reactive_power,
    'Apparent Power (MVA)': apparent_power,
    'Temperature (°C)': temperature,
    'Humidity (%)': humidity,
    'Wind Speed (m/s)': wind_speed,
    'Solar Radiation (W/m²)': solar_radiation,
    'Load (% Capacity)': load_capacity,
    'Line Load (%)': line_load,
    'Insulation Resistance (MΩ)': insulation_resistance,
    'Leakage Current (mA)': leakage_current
})

# Write to Excel
df.to_excel('power_grid_data.xlsx', index=False)

# Read the Excel file
df = pd.read_excel('power_grid_data.xlsx')

# Print the first 5 rows of the DataFrame
print(df.head())

# Convert 'Timestamp' to datetime
df['Timestamp'] = pd.to_datetime(df['Timestamp'])

# Plot 'Voltage (kV)' over time
plt.figure(figsize=(10, 6))
plt.plot(df['Timestamp'], df['Voltage (kV)'])
plt.title('Voltage over Time')
plt.xlabel('Time')
plt.ylabel('Voltage (kV)')
plt.show()

# Calculate the average temperature
average_temperature = df['Temperature (°C)'].mean()

print(f"The average temperature is {average_temperature} °C")

# Calculate the maximum humidity
max_humidity = df['Humidity (%)'].max()

print(f"The maximum humidity is {max_humidity} %")

# Calculate the minimum power factor
min_power_factor = df['Power Factor'].min()

print(f"The minimum power factor is {min_power_factor}")

# Calculate the total active power
total_active_power = df['Active Power (MW)'].sum()

print(f"The total active power is {total_active_power} MW")

