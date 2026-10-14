"""
Dataset Generator: Official UPES Dehradun College Bus Telemetry
Routes:
  - Bus 7717: Miyawala Express (via Jogiwala, Rispana)
  - Bus 7267: Miyawala - ISBT - FRI City Loop (17 stops)
  - Bus 0303 / 8527 / 1757: Jakhan (Rajpur Road Corridor)
  - Bus 6727 / 9247: Raipur Corridor
  - Bus 8437: IT Park Corridor
Saves to 'data/upes_bidholi_commute_raw.csv'
"""

import os
import numpy as np
import pandas as pd

np.random.seed(42)
n_samples = 3000

# 1. Official Bus & Route Profiles from UPES Transport Division
bus_routes_data = [
    {'bus_no': '7717', 'route_name': 'Miyawala_Express', 'driver': 'Sudhir', 'base_dist_km': 34.0, 'base_scheduled_min': 85},
    {'bus_no': '7267', 'route_name': 'Miyawala_ISBT_FRI_Loop', 'driver': 'Bahadur_Singh', 'base_dist_km': 39.0, 'base_scheduled_min': 80},
    {'bus_no': '0303', 'route_name': 'Jakhan_Rajpur_Road', 'driver': 'Sunny', 'base_dist_km': 28.0, 'base_scheduled_min': 70},
    {'bus_no': '6727', 'route_name': 'Raipur_Corridor', 'driver': 'Pradeep', 'base_dist_km': 32.0, 'base_scheduled_min': 75},
    {'bus_no': '8437', 'route_name': 'IT_Park_Corridor', 'driver': 'Monu', 'base_dist_km': 29.0, 'base_scheduled_min': 70},
    {'bus_no': '5860', 'route_name': 'Chandra_Mani_Route', 'driver': 'Bittu', 'base_dist_km': 26.0, 'base_scheduled_min': 65}
]

# Randomly assign trips among the real bus fleet
chosen_buses = np.random.choice(bus_routes_data, size=n_samples)

bus_numbers = [b['bus_no'] for b in chosen_buses]
route_names = [b['route_name'] for b in chosen_buses]
driver_names = [b['driver'] for b in chosen_buses]
base_distances = np.array([b['base_dist_km'] for b in chosen_buses])
base_scheduled_times = np.array([b['base_scheduled_min'] for b in chosen_buses])

# Actual measured GPS distance with slight deviations (detours / diversions)
actual_distance_km = base_distances + np.random.uniform(-1.0, 2.5, size=n_samples)

# Destination Campuses (Bidholi vs Kandoli)
dest_campus = np.random.choice(['Bidholi_Campus', 'Kandoli_Campus'], size=n_samples, p=[0.75, 0.25])

# Departure slots
departure_slot = np.random.choice(['Slot_6_15_AM', 'Slot_7_00_AM', 'Slot_2_00_PM', 'Slot_5_30_PM'], 
                                  size=n_samples, p=[0.25, 0.45, 0.10, 0.20])

# Real environmental factors on the Prem Nagar -> Nanda Ki Chowki -> Bidholi Ghat road
monsoon_rain_intensity_mm = np.random.exponential(scale=5.0, size=n_samples)
nanda_ki_chowki_chokepoints = np.random.poisson(lam=4, size=n_samples)
sensor_noise = np.random.normal(loc=1.0, scale=0.005, size=n_samples)

# Introduce realistic flaws for Unit IV Preprocessing demo:
# 1. Hill GPS signal drop (missing distance)
actual_distance_km[0:25] = np.nan
# 2. Extreme telemetry glitch
nanda_ki_chowki_chokepoints[40:45] = 28

# Calculate Actual Travel Time in Minutes
time_multiplier = np.where(departure_slot == 'Slot_7_00_AM', 35.0,
                  np.where(departure_slot == 'Slot_5_30_PM', 25.0,
                  np.where(departure_slot == 'Slot_6_15_AM', 5.0, 0.0)))

campus_extra = np.where(dest_campus == 'Bidholi_Campus', 12.0, 0.0)

# Clean distance for realistic ground truth calculation
clean_dist = np.nan_to_num(actual_distance_km, nan=30.0)

# Delay calculation: Scheduled duration is already baseline travel time.
# Congestion added on top: slot delay + rain + choke points
congestion_delay_mins = (time_multiplier + campus_extra + 
                         (np.clip(nanda_ki_chowki_chokepoints, 0, 10) * 2.5) + 
                         (monsoon_rain_intensity_mm * 1.2) + 
                         (clean_dist - base_distances) * 1.5 +
                         np.random.normal(0, 4, size=n_samples))

actual_travel_time_mins = base_scheduled_times + np.maximum(0, congestion_delay_mins)

# Classification Target: Severe Delay Alert (> 65 mins delay beyond scheduled time - ~10-15% imbalanced)
delay_minutes = actual_travel_time_mins - base_scheduled_times
severe_delay_alert = (delay_minutes > 65.0).astype(int)

# Create DataFrame
df = pd.DataFrame({
    'Trip_ID': [f"UPES_{b}_{1000+i}" for i, b in enumerate(bus_numbers)],
    'Bus_Number': bus_numbers,
    'Route_Corridor': route_names,
    'Driver_Name': driver_names,
    'Destination_Campus': dest_campus,
    'Departure_Slot': departure_slot,
    'Scheduled_Duration_Min': base_scheduled_times,
    'GPS_Distance_KM': actual_distance_km,
    'Monsoon_Rain_MM': monsoon_rain_intensity_mm,
    'Ghat_Choke_Points': nanda_ki_chowki_chokepoints,
    'Bus_Sensor_Noise': sensor_noise,
    'Actual_Travel_Time_Min': actual_travel_time_mins,
    'Severe_Delay_Alert': severe_delay_alert
})

# Save to data directory and root for convenience
base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
data_dir = os.path.join(base_dir, 'data')
os.makedirs(data_dir, exist_ok=True)

csv_file = os.path.join(data_dir, "upes_bidholi_commute_raw.csv")
df.to_csv(csv_file, index=False)

# Also save directly in project root if executed from root
root_csv = os.path.join(base_dir, "upes_bidholi_commute_raw.csv")
df.to_csv(root_csv, index=False)

print("="*70)
print(f"[SUCCESS] UPES Transit Dataset Created: '{csv_file}'")
print(f"Total Logged Trips: {df.shape[0]} | Columns: {df.shape[1]}")
print("Buses Tracked:", df['Bus_Number'].unique())
print("Routes Tracked:", df['Route_Corridor'].unique())
print("Severe Delay Rate:", f"{df['Severe_Delay_Alert'].mean()*100:.2f}%")
print("="*70)
