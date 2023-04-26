import pandas as pd
from geopy.geocoders import Nominatim

# read the data from the Excel sheet
data = pd.read_excel("production_plant_data.xlsx", sheet_name="COGLOCATION")

# use geopy to look up the coordinates and zip codes of the locations, but skip rows that already have coordinates and zip codes
geolocator = Nominatim(user_agent="production_plant")
for index, row in data.iterrows():
    if pd.isna(row["latitude"]) or pd.isna(row["longitude"]) or pd.isna(row["zip"]):
        location = geolocator.geocode(row["location"], addressdetails=True)
        if location is not None:
            data.at[index, "latitude"] = location.latitude
            data.at[index, "longitude"] = location.longitude
            if "postcode" in location.raw["address"]:
                data.at[index, "zip"] = location.raw["address"]["postcode"]

# calculate the total weight of all objects
total_weight = data["weight"].sum()

# calculate the x and y coordinates of the center of gravity
x_cg = (data["longitude"] * data["weight"]).sum() / total_weight
y_cg = (data["latitude"] * data["weight"]).sum() / total_weight

# use geopy to look up the address of the center of gravity
location = geolocator.reverse("{}, {}".format(y_cg, x_cg))

# print the center of gravity and location
print("The center of gravity is at longitude = {:.2f}, latitude = {:.2f}".format(x_cg, y_cg))
print("The center of gravity is located at {}".format(location.address))