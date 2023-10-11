map_plot = folium.Map(location=[data['latitude'].mean(), data['longitude'].mean()], zoom_starts=4, tiles='CartoDB positron')
    country_geojson = country_geometry_high_res.__geo_interface__
    folium.GeoJson(country_geojson, name='Country Boundary',style_function=lambda feature:{
                        'fillColor': 'lightyellow',
                        'color': 'black',
                        'weight': 1,
                        'fillOpacity': 0.3,      
                        }
                      ).add_to(map_plot)
    
    heat_data = [[row['latitude'], row['longitude']] for index, row in data.iterrows() if Point(row['longitude'], row['latitude']).within(country_geometry_low_res)]
    marker_cluster = folium.plugins.MarkerCluster(name="Coordinate Points", options={'disableClusteringAtZoom':10},show=False).add_to(map_plot)

    #overlay group
    overlay_group = folium.FeatureGroup(name='Overlay Layers',show=False)
    overlay_group.add_to(map_plot)

    if user_selection == 'H':
        HeatMap(heat_data,name="HeatMap", show = False).add_to(map_plot)

    elif user_selection == 'HO':
        HeatMap(heat_data,name="HeatMap", show = False).add_to(map_plot)
        for lat_long,count in inside_count.items():
            folium.Circle(lat_long,popup={lat_long},
                            radius=count*5, color ='black', fill=True,fill_opacity=0.5).add_to(overlay_group)

    elif user_selection == 'P':
        for lat_long,count in inside_count.items():
            folium.Marker(lat_long,popup=f"Inside Boundaries<br>{lat_long}",
                            icon=folium.Icon(color='green')).add_to(marker_cluster)
            folium.Circle(lat_long,popup={lat_long},
                            radius=count*5, color ='black', fill=True, fill_opacity=0.5).add_to(overlay_group)
    else:
        HeatMap(heat_data, name="HeatMap", show = False).add_to(map_plot)
        for lat_long,count in inside_count.items():
            folium.Marker(lat_long,popup=f"Inside Boundaries<br>{lat_long}",
                            icon=folium.Icon(color='green')).add_to(marker_cluster)
            folium.Circle(lat_long,popup={lat_long},
                            radius=count*5, color ='black', fill=True, fill_opacity=0.5).add_to(overlay_group)

    for ID,lat,lon in outside_high_res:
        folium.Marker(location=[lat,lon], popup=f"Outside Boudnaries<br>ID:{ID}<br>{lat}<br>{lon}",
                            icon=folium.Icon(color='red')).add_to(map_plot)

    #Adding legend for differentiating Inside and Outside coordinate points
    legend_html = """
        <div style="position: fixed; bottom:30px; right:30px; z-index:9999; font-size:11px;">
            <p><i class="fa fa-circle fa-1x" style="color:green"></i> Inside Boundary</p>
            <p><i class="fa fa-circle fa-1x" style="color:red"></i> Outside Boundary</p>
        </div>
        """
    map_plot.get_root().html.add_child(folium.Element(legend_html))

    #Adding various tiles option incoorporated within Folium library
    folium.TileLayer('CartoDB positron').add_to(map_plot) #Selected as Default
    folium.TileLayer('OpenStreetMap').add_to(map_plot) 
    folium.TileLayer('Stamen Terrain').add_to(map_plot)
    folium.TileLayer('Stamen Toner').add_to(map_plot)
    folium.TileLayer('esrinatgeoworldmap', name='Esri NatGeo WorldMap', attr=' Esri, Delorme, NAVTEQ').add_to(map_plot)

    #Full Screen option
    folium.plugins.Fullscreen().add_to(map_plot)

    #Layer control for each layers added can be checked & unchecked while selecting
    folium.LayerControl(position='topright',collapsed=True, autoZIndex=True).add_to(map_plot)

    #Saved as map.html 
    map_filename = excel_doc.replace('.xlsx', '.html')
    map_plot.save(map_filename)
    
    print(f"Total number of Coordinates in Excel file are : {total_coord}")
    print(f"Map created and no. of Coord falling Inside {user_input} are: {len(inside)} and Outside {user_input} are: {len(outside_high_res)}")
