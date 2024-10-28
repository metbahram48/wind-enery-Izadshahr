import numpy as np
import pandas as pd
from scipy.stats import weibull_min
import matplotlib.pyplot as plt
import calendar

# Load the data from Excel file
df = pd.read_excel(r'C:\Users\TAK\Desktop\project\wind IZAD.xlsx')

# Calculate effective wind speed
data['effective_wind_speed'] = (w_avg * data['avg_wind_speed'] +
                                w_max * data['max_wind_speed'])/2

# Adjust wind speed for 50m height using the 1/7 power law
data['adjusted_effective_wind_speed'] = data['effective_wind_speed'] * (50 / 10) ** (1/7)

# Fit Weibull distribution to effective wind speed at 50m height
shape, loc, scale = weibull_min.fit(data['adjusted_effective_wind_speed'], floc=0)

# Calculate power density (kW/m²) based on effective wind speed at 50m height
air_density = 1.225  # kg/m³, at sea level
data['wind_power_density_kw'] = 0.5 * air_density * (data['adjusted_effective_wind_speed'] ** 3) / 1000  # in kW/m²

# Monthly and yearly wind power density aggregation
data['month'] = pd.to_datetime(data.index).month
data['month_name'] = data['month'].apply(lambda x: calendar.month_name[x])

# Sum power density per month to get energy in kWh
hours_per_day = 24
days_in_month = data.groupby('month')['wind_power_density_kw'].count()
monthly_energy_kwh = (data.groupby('month')['wind_power_density_kw'].sum() * days_in_month * hours_per_day)
yearly_energy_kwh = monthly_energy_kwh.sum()

# Plot Monthly Wind Power Density in kW/m² with month names
plt.figure(figsize=(10, 6))
monthly_energy_kwh.index = monthly_energy_kwh.index.map(lambda x: calendar.month_name[x])
monthly_energy_kwh.plot(kind='bar', color='skyblue')
plt.title("Monthly Wind Power Density (kW/m²)")
plt.xlabel("Month")
plt.ylabel("Wind Power Density (kW/m²)")
plt.xticks(rotation=45)
plt.show()

# Plot Weibull Distribution Fit for Effective Wind Speed at 50m
x = np.linspace(0, data['adjusted_effective_wind_speed'].max(), 100)
weibull_pdf = weibull_min.pdf(x, shape, loc=loc, scale=scale)

plt.figure(figsize=(10, 6))
plt.hist(data['adjusted_effective_wind_speed'], bins=20, density=True, alpha=0.6, color='skyblue', edgecolor='black')
plt.plot(x, weibull_pdf, 'r-', lw=2, label=f'Weibull fit\nShape={shape:.2f}, Scale={scale:.2f}')
plt.title("Effective Wind Speed Distribution at 50m and Weibull Fit")
plt.xlabel("Effective Wind Speed (m/s)")
plt.ylabel("Density")
plt.legend()
plt.show()

# Wind Class Determination at 50m Height
mean_wind_speed_50m = data['adjusted_effective_wind_speed'].mean()
if mean_wind_speed_50m < 4.4:
    wind_class = "Class 1 (Poor)"
elif 4.4 <= mean_wind_speed_50m < 5.4:
    wind_class = "Class 2 (Marginal)"
elif 5.4 <= mean_wind_speed_50m < 6.7:
    wind_class = "Class 3 (Good)"
elif 6.7 <= mean_wind_speed_50m < 8.0:
    wind_class = "Class 4 (Very Good)"
else:
    wind_class = "Class 5 (Excellent)"

# Display results
print("Weibull Shape (k) and Scale (c) Parameters:", shape, scale)
print("Monthly Energy Produced (kWh):\n", monthly_energy_kwh)
print("Yearly Energy Produced (kWh):", yearly_energy_kwh)
print("Wind Class at 50m:", wind_class)
