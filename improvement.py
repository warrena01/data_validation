'''
Changes to make:
-> make the excel data document a gdf and ensure that they have the same crs
-> add analysis per crester zone (if it is within the crester zone, and if it is in the centroid)
-> change column names on the original data import to ensure they have the correct capitalisation

-> create a buffer around the original low_res_geometry

Changes made:
-> automatically goes through at a higher resolution, but for the smaller amount of points
-> improved efficiency for the counter dictionary 
'''
try:
    import tkinter as tk
    from tkinter import filedialog
    import geopandas as gpd
    import pandas as pd
    import folium
    from shapely.geometry import Point
    root = tk.Tk()
    from folium.plugins import HeatMap, MarkerCluster

    '''GET THE DATA'''

    # get the excel file that contains the coordinates
    excel_doc = filedialog.askopenfilename()
    root.withdraw()
    data = pd.read_excel(excel_doc)
    data.columns = ['ID', 'latitude', 'longitude']

    # get the shapefile for the area(s)
    shapefile_low_res = filedialog.askopenfilename()
    root.withdraw()
    gdf_low_res = gpd.read_file(shapefile_low_res)

    # get the shapefile for the area(s)
    shapefile_high_res = filedialog.askopenfilename()
    root.withdraw()
    root.destroy()
    gdf_high_res = gpd.read_file(shapefile_high_res)


    '''GEOSPATIAL ANALYSIS'''

    # choose country of analysis and get geometry
    user_input = input('input the ISO_A3 for the country you are analysing, the hit enter:  ').upper()
    while user_input not in gdf['ISO_A3'].tolist():
        user_input = input('You have inputted an invalid ISO A3 Value, please try again:  ').upper()
    country_gdf_low_res = gdf_low_res.loc[gdf_low_res['ISO_A3'] == user_input]
    country_geometry_low_res = country_gdf_low_res.geometry.iloc[0]
    country_gdf_high_res = gdf_high_res.loc[gdf_high_res['ISO_A3'] == user_input]
    country_geometry_low_res = country_gdf_low_res.geometry.iloc[0]

    inside = []
    outside_low_res = pd.DataFrame(columns=user_input.columns)
    outside_high_res = []
    inside_count = {}
    outside_count = {}

    # points within the country analysis
    for id, row in data.iterrows():
        location = Point(row['latitiude'], row['longitude'])
        if point.within(country_geometry_low_res):
            inside.append((row['ID'], row['latitude'], row['longitude']))
            # update counter dictionary
            lat_long = (row['latitude'], row['longitude']) 
            inside_count[lat_long] = inside_count.get(lat_long, 0)+1
        else:
            gdf_temporary = geopandas.GeoDataFrame([row], columns=outside_low_res.columns, crs=outside_low_res.crs)
            outside_low_res = pd.concat([outside_low_res, gdf], ignore_index=True)
    for id, row in outside_low_res.iterrows():
        location = Point(row['latitiude'], row['longitude'])
        if point.within(country_geometry_high_res):
            inside.append((row['ID'], row['latitude'], row['longitude']))
            # update counter dictionary
            lat_long = (row['latitude'], row['longitude']) 
            inside_count[lat_long] = inside_count.get(lat_long, 0)+1
        else:
            outside_high_res.append((row['ID'], row['latitude'], row['longitude']))

    '''CREATE MAP OUTPUTS'''

    # create base map
    base_map = folium.Map(location = [data['latitude'].mean(), data['longitude'].mean()], zoom_starts=3, tiles='CartoDB positron')
    # add base layer of the country
    country_geojson = country_geometry_low_res.__geo_interface__
    folium.GeoJson(country_geojson, name = 'Country Border', style_function = lambda feature:{
        'fillColor': 'lightyellow',
        'color': 'black',
        'weight': 1,
        'fillOpacity': 0.3}
        ).add_to(base_map)
    
    
    # add a legend
    legend_html = """
                    <div style="position: fixed; bottom:30px; right:30px; z-index:9999; font-size:11px;">
                        <p><i class"fa fa-circle fa-1x" style="color:green">Within Boundary</p>
                        <p><i class"fa fa-circle fa-1x" style="color:red">Outisde Boundary</p>
                    </div>
                  """
    base_map.get_root().html.add_child(folium.Element(legend_html))
    
    # add the various map elements
    folium.TileLayer('CartoDB positron').add_to(base_map) 
    folium.TileLayer('OpenStreetMap').add_to(base_map) 
    folium.TileLayer('Stamen Terrain').add_to(base_map) 
    folium.TileLayer('Stamen Toner').add_to(base_map) 
    folium.TileLayer('esrinatgeoworldmap', name='Esri NatGeo WorldMap', attr=' Esri, Delorme, NAVTEQ').add_to(base_map) 

    # full screen option
    folium.plugins.Fullscreen().add_to(base_map)
    
    # layer control for user to control
    folium.LayerControl(position='topright', collapsed=True, autoZIndex=True).add_to(base_map)
    
    
    
except ImportError:
    raise ImportError("The required library is not installed")
