from rtree.index import Index
import geopandas as gpd
import pandas as pd
from shapely.geometry import Point 
import folium
from folium.plugins import MarkerCluster
import tkinter
from tkinter import filedialog

'''ORGANISE INPUTS'''

data = None
gdf = None
dir_path = None

def get_excel_doc():
    global data 
    root = tkinter.Tk()
    excel_doc = filedialog.askopenfilename()
    root.withdraw()
    if excel_doc.split(".")[-1] in ["xlsx", "xls"]:
        data = pd.read_excel(excel_doc)
    else:
        print("you chose the wrong file type, try again.")
get_excel_doc()

def get_gdf():
    global gdf
    root = tkinter.Tk()
    my_file_path = filedialog.askopenfilename()
    root.withdraw()
    if my_file_path.split(".")[-1] in ["shp"]:
        gdf = gpd.read_file(my_file_path)
    else:
        print("you chose the wrong file type, try again.")
get_gdf()

def get_file_directory():
    global dir_path
    root = tkinter.Tk()
    dir_path = filedialog.askdirectory()
    root.withdraw()       
get_file_directory()

def create_folium(inside_df, inside_points_dict, ambiguity_points_dict, outside_points_dict, file_path_to_save):
    # Create a Folium map centered around the mean
    m = folium.Map(location=[inside_df['LAT'].mean(), inside_df['LONG'].mean()], zoom_start=4)

    # Add markers for points inside the cresta
    inside_cluster = MarkerCluster(name="Correctly Geocoded Coordinate Points", options={'disableClusteringAtZoom':10},show=False).add_to(m)
    for (lat_long, polygon), info in inside_points_dict.items():
        count = info['count']
        folium.Marker(lat_long, popup=f'''<h5 style="width: 120px;">Inside</h5><p style="width: 120px;">Cresta: {polygon}<br>Count: {count}<br>Latitude: {lat_long[0]}<br>Longitude: {lat_long[1]}</p>''',
                      max_width=1500, icon=folium.Icon(color='green')).add_to(inside_cluster)

    # Add markers for points with cresta ambiguity
    cresta_ambiguity_cluster = MarkerCluster(name="Cresta/Coordinate Inconsistency", options={'disableClusteringAtZoom':10},show=True).add_to(m)
    for (lat_long, polygon_stated), info in ambiguity_points_dict.items():
        count = info['count']
        polygon_found = info.get('polygon_found', 'Not Found')
        folium.Marker(lat_long, 
                      popup=f'''<h5 style="width: 120px;">Ambiguity</h5><p style="width: 120px;">Cresta Stated: {polygon_stated}, Cresta Found: {polygon_found}<br>Count: {count}<br>Latitude: {lat_long[0]}<br>Longitude: {lat_long[1]}</p>''',
                      max_width=1500, icon=folium.Icon(color='orange')).add_to(cresta_ambiguity_cluster)

    # Add markers for points outside the cresta
    outside_cluster = MarkerCluster(name="Outside Cresta Zones", options={'disableClusteringAtZoom':10},show=True).add_to(m)
    for (lat_long, polygon_stated), info in outside_points_dict.items():
        count = info['count']
        folium.Marker(lat_long, 
                      popup=f'''<h5 style="width: 120px;">Outside</h5><p style="width: 120px;">Cresta Stated: {polygon_stated}<br>Count: {count}<br>Latitude: {lat_long[0]}<br>Longitude: {lat_long[1]}</p>''',
                      max_width=1500, icon=folium.Icon(color='red')).add_to(outside_cluster)

    # Add the crestas on to the map as individual zones and as a whole
    all_crestas = folium.FeatureGroup(name='All Crestas', show=True)
    for idx, row in country_gdf.iterrows():
        folium.GeoJson(row['geometry']).add_to(all_crestas)
    all_crestas.add_to(m)   
    for idx, row in country_gdf.iterrows():
        polygon_group = folium.FeatureGroup(name=row['CRESTA_DES'], show=False)
        folium.GeoJson(row['geometry']).add_to(polygon_group)
        polygon_group.add_to(m)

    # Adding various tiles option incorporated within Folium library
    folium.TileLayer('CartoDB positron').add_to(m) #Selected as Default
    folium.TileLayer('OpenStreetMap').add_to(m) 

    # get map output
    m.save(file_path_to_save) 

def CamelCase(word):
    words = word.split() 
    camel_case = ' '.join(word.capitalize() for word in words)
    return camel_case

# user will define the country such that the gdf only contains relevant information.
country_gdf = gdf.loc[gdf['COUNTRY'] == 'Peru']

# start a spatial index
my_spatial_index = Index()
for id, country in country_gdf.iterrows():
    my_spatial_index.insert(id, country.geometry.bounds)

# initialise variables
inside = {}
inside_id = []
cresta_ambiguity = {}
cresta_ambiguity_id = []
outside = {}
outside_id = []

# check each row of data
for idx, row in data.iterrows():
    point = Point(row['LONG'], row['LAT']) 
    lat_long = (row['LAT'], row['LONG']) 
    # find the cresta the client has given, standardize input
    client_cresta = CamelCase(row['REGION'])   
    # get the polygon for given cresta
    cresta_geometry = (country_gdf.loc[country_gdf['CRESTA_DES'] == client_cresta, 'geometry']).values[0]
    
    if point.within(cresta_geometry):
        key = (lat_long, client_cresta)
        inside[key] = inside.get(key, {'polygon': client_cresta, 'count': 0})
        inside[key]['count'] += 1
        inside_id.append(row['LOCNUM'])
    
    else: 
        possible_matches_id = list(my_spatial_index.intersection(point.bounds))
        if possible_matches_id:
            for row_id in possible_matches_id:
                if country_gdf.geometry[row_id].contains(point):
                    key = (lat_long, client_cresta)
                    cresta_ambiguity[key] = cresta_ambiguity.get(key, {'polygon_stated': client_cresta, 'polygon_found': country_gdf['CRESTA_DES'][row_id], 'count': 0})
                    cresta_ambiguity[key]['count'] += 1
                    cresta_ambiguity_id.append(row['LOCNUM'])
                else: 
                    continue      

        else:
            key = (lat_long, client_cresta)
            outside[key] = outside.get(key, {'polygon_stated': client_cresta, 'count': 0})
            outside[key]['count'] += 1
            outside_id.append(row['LOCNUM'])
            
'''CREATE OUTPUTS '''

# create  outputs ready for excel
inside_df = data[data['LOCNUM'].isin(inside_id)]
inside_df_path = dir_path + (r'/Correct Cresta Zone and Lat Longs.xlsx')
inside_df.to_excel(inside_df_path, index=False)

cresta_ambiguity_df = data[data['LOCNUM'].isin(cresta_ambiguity_id)]
cresta_ambiguity_df_path = dir_path + (r'/Inconsistent Cresta Zone and Lat Longs.xlsx')
cresta_ambiguity_df.to_excel(cresta_ambiguity_df_path, index=False)

outside_df = data[data['LOCNUM'].isin(outside_id)]
outside_df_path = dir_path + (r'/Outside Any Cresta Zone Lat Longs.xlsx')
outside_df.to_excel(outside_df_path, index=False)

map_path = dir_path + r'/Map Output.html'
create_folium(inside_df, inside, cresta_ambiguity, outside, map_path)

# user reporting text file
percentage_ad = len(cresta_ambiguity_df)*100/len(data)
percentage_ocz = len(outside_df)*100/len(data)

user_reporting_file_path = dir_path + (r'/User Reporting.txt')
file = open(user_reporting_file_path, "w")
file.write(f'Total correct cresta data points: {len(inside_id)} \n'
            f'Total data points with inconsistent cresta/latlongs: {len(cresta_ambiguity_df)} \n'
            f'Total data points outside any cresta zone: {len(outside_id)} \n'
            f'Total points: {len(data)} \n'
            f'Percentage of ambiguous data: {percentage_ad:.1f}% \n'
            f'Percentage of data outside any cresta zone: {percentage_ocz:.1f}% \n' )
file.close()
        
