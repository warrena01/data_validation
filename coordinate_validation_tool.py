"""
Created on Thu Dec  7 16:30:02 2023

@author: adwarren
"""
"""
improvements in v3:

- all 4 apps now available in 1 app
- speed improvements when exporting excel files via .csv function
- speed improvements via dict used to check for repetitions to reduce spatial checks performed
- speed improvements via cresta check only using inside_ll data if ran subsequent to country check
- first 1000 rows option integrated to allow for mistakes to be spotted more efficiently
- speed improvements in geolocator via dictionary 

"""
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QLabel, QFileDialog, QComboBox, QTextEdit,QProgressBar
from PyQt5.QtCore import Qt
import sys
import pandas as pd
import os
import glob
import geopandas as gpd
import shapely
from geopy.geocoders import Nominatim
from rtree.index import Index
import numpy
import folium
from folium.plugins import MarkerCluster
from folium.plugins import HeatMap
import time
import numpy as np


class CoordinateValidationTool(QMainWindow):
    
    def __init__(self):
        super().__init__()

        self.country_name = None
        self.first_country_name = None
        self.analysis_type = None
        self.excel_path = None
        self.data = None
        self.loc_col = None
        self.lat_col = None
        self.long_col = None
        self.cresta_col = None
        
        self.nan_data = None
        self.inside_df_country_check = None
        self.outside_df_country_check = None
        self.outside_df_country_reversed_checked = None
        self.inside_df_cresta_check = None
        self.outside_df_cresta_check = None
        self.ambiguous_df_cresta_check = None
        self.country_inside_dict = None
        self.country_outside_dict = None
        self.cresta_inside_dict = None
        self.cresta_outside_dict = None
        self.cresta_ambiguous_dict = None
        
        self.first_1000_idx = None
        self.cresta_just_run = None
        self.analysis_just_run = False
        self.recent_analysis = None

        self.progress_bar = QProgressBar(self)
        self.progress_bar.setGeometry(400, 410, 250, 20)
        self.progress_bar.setOrientation(Qt.Horizontal)
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0) 
        
        self.initUI()
        
    def initUI(self):

        # create the title and the geometry of the window
        self.setWindowTitle('GeoData Validation Tool')
        self.setGeometry(250, 250, 650, 800)
        
        ''' Get Excel '''
        self.label1 = QLabel('Step 1: Select the excel workbook', self)
        self.label1.setGeometry(30, 20, 400, 20)
        self.btn1 = QPushButton('Choose File', self)
        self.btn1.setGeometry(30, 50, 100, 30)
        self.btn1.clicked.connect(self.get_excel_doc)
        
        ''' Decide which geospatial analysis '''
        self.label2 = QLabel('Step 2: Decide Which Spatial Analysis To Do', self)
        self.label2.setGeometry(30, 100, 400, 20)
        self.listbox = QComboBox(self)
        self.listbox.setGeometry(30, 130, 250, 30)
        self.listbox.addItems(['Country First 1000 Rows', 'Country', 'Cresta'])
        self.btn2 = QPushButton('Confirm', self)
        self.btn2.setGeometry(30, 160, 100, 30)
        self.btn2.clicked.connect(self.choose_analysis)
        
        ''' Decide which geospatial analysis '''
        self.label3 = QLabel('Step 3: Choose Country of Analysis', self)
        self.label3.setGeometry(30, 210, 400, 20)
        self.listbox2 = QComboBox(self)
        self.listbox2.setGeometry(30, 240, 250, 30)
        self.listbox2.addItems([])
        self.btn3 = QPushButton('Confirm', self)
        self.btn3.setGeometry(30, 270, 100, 30)
        self.btn3.clicked.connect(self.choose_country)
        
        ''' Map the relevant columns '''
        self.label4 = QLabel('Step 4: Map your columns', self)
        self.label4.setGeometry(30, 320, 500, 20)
        
        self.label5 = QLabel('Locnum', self)
        self.label5.setGeometry(30, 350, 100, 30)
        self.listbox3 = QComboBox(self)
        self.listbox3.setGeometry(150, 350, 200, 30)
        self.listbox3.addItems([""])
        self.listbox3.setCurrentIndex(-1)

        self.label6 = QLabel('Latitude', self)
        self.label6.setGeometry(30, 380, 100, 30)
        self.listbox4 = QComboBox(self)
        self.listbox4.setGeometry(150, 380, 200, 30)
        self.listbox4.addItems([""])
        self.listbox4.setCurrentIndex(-1)

        self.label7 = QLabel('Longitude', self)
        self.label7.setGeometry(30, 410, 100, 30)
        self.listbox5 = QComboBox(self)
        self.listbox5.setGeometry(150, 410, 200, 30)
        self.listbox5.addItems([""])
        self.listbox5.setCurrentIndex(-1)
        
        self.label7 = QLabel('Cresta', self)
        self.label7.setGeometry(30, 440, 100, 30)
        self.listbox6 = QComboBox(self)
        self.listbox6.setGeometry(150, 440, 200, 30)
        self.listbox6.addItems([""])
        self.listbox6.setCurrentIndex(-1)
        self.label7.setEnabled(False)
        self.listbox6.setEnabled(False)

        self.btn4 = QPushButton('Confirm', self)
        self.btn4.setGeometry(150, 470, 100, 30)
        self.btn4.clicked.connect(self.standardize_cols)
        
        ''' Outputs '''
        self.label8 = QLabel('Step 5: Start Analysis', self)
        self.label8.setGeometry(30, 520, 300, 20)

        self.btn5 = QPushButton('Go', self)
        self.btn5.setGeometry(30, 550, 100, 30)
        self.btn5.clicked.connect(self.geospatial_analysis)
        
        self.reverse_geocode_btn = QPushButton('Run Reverse Geocode', self)
        self.reverse_geocode_btn.setGeometry(30, 580, 200, 30)
        self.reverse_geocode_btn.clicked.connect(self.reverse_geocode)
        
        self.map_output_button = QPushButton('Produce Map', self)
        self.map_output_button.setGeometry(30, 610, 200, 30)
        self.map_output_button.clicked.connect(self.get_map_output) 
        self.map_output_button.setEnabled(False)      
        
        # Text Box
        self.textOutput = QTextEdit(self)
        self.textOutput.setGeometry(400, 30, 230, 380)
        self.textOutput.setReadOnly(True)
        
        self.show()
        
        self.country_name = None
        self.analysis_type = None
        self.excel_path = None
        self.data = None
        self.loc_col = None
        self.lat_col = None
        self.long_col = None
        self.cresta_col = None
        self.nan_data = None
        
    def get_map_output(self):

        try:

            if len(self.data) > 0:

                self.update_progress(1)

                m = folium.Map(location = [self.data['LAT_AJG'].mean(), self.data['LONG_AJG'].mean()], zoom_start=4, control_scale=True, tiles='CartoDB positron')
                dir_path = self.excel_path.replace(self.excel_path.split("/")[-1], "")
                file_save_path = dir_path + f'\Map Output {self.analysis_type}.html'

                if self.recent_analysis == 'Country':

                    country_geojson = self.country_geometry_high_res.__geo_interface__
                    folium.GeoJson(country_geojson, name='Country Boundary',style_function=lambda feature:{
                                'fillColor': 'lightyellow',
                                'color': 'black',
                                'weight': 1,
                                'fillOpacity': 0.3,      
                                }
                            ).add_to(m)

                    heat_data = [[row['LAT_AJG'], row['LONG_AJG']] for index, row in self.inside_df_country_check.iterrows()]
                    HeatMap(heat_data,name="HeatMap", show = False).add_to(m)
                    
                    marker_cluster = MarkerCluster(name="Outside Points", options={'disableClusteringAtZoom':10},show=False).add_to(m)
                    for lat_long,count in self.country_outside_dict.items():
                        folium.Marker(lat_long,popup=f"Invalid Coordinate<br>{lat_long}<br>Count: {count}",
                                        icon=folium.Icon(color='red')).add_to(marker_cluster)
                    
                     # Adding legend for differentiating Inside and Outside coordinate points
                    legend_html = """
                        <div style="position: fixed; bottom:30px; right:30px; z-index:9999; font-size:11px;">
                            <p><i class="fa fa-circle fa-1x" style="color:red"></i> Outside Country </p>
                        </div>
                        """
                    
                    # Adding various tiles option incorporated within Folium library
                    folium.TileLayer('CartoDB positron').add_to(m) # Selected as Default
                    folium.TileLayer('OpenStreetMap').add_to(m) 

                    # Layer control for each layers added can be checked & unchecked while selecting
                    folium.LayerControl(position='topright',collapsed=True, autoZIndex=True).add_to(m)

                    m.get_root().html.add_child(folium.Element(legend_html))

                    # Full Screen option
                    folium.plugins.Fullscreen().add_to(m)

                    m.save(file_save_path)

                    self.append_text('Map Output Successfully Saved')
                    self.update_progress(100)

                if self.recent_analysis == 'Cresta':

                    # Add markers for points with cresta ambiguity
                    cresta_ambiguity_cluster = MarkerCluster(name="Cresta/Coordinate Inconsistency", options={'disableClusteringAtZoom':10},show=True).add_to(m)
                    for (lat_long, polygon_stated), info in self.cresta_ambiguous_dict.items():
                        count = info['count']
                        polygon_found = info.get('polygon_found', 'Not Found')
                        folium.Marker(lat_long, 
                                    popup=f'''<h5 style="width: 120px;">Ambiguity</h5><p style="width: 120px;">Cresta Stated: {polygon_stated}<br> Cresta Found: {polygon_found}<br>Count: {count}<br>Latitude: {lat_long[0]}<br>Longitude: {lat_long[1]}</p>''',
                                    max_width=1500, icon=folium.Icon(color='orange')).add_to(cresta_ambiguity_cluster)
                    
                    # Add markers for points outside the cresta
                    outside_cluster = MarkerCluster(name="Outside Cresta Zones", options={'disableClusteringAtZoom':10},show=True).add_to(m)
                    for (lat_long, polygon_stated), info in self.cresta_outside_dict.items():
                        count = info['count']
                        folium.Marker(lat_long, 
                                    popup=f'''<h5 style="width: 120px;">Outside</h5><p style="width: 120px;">Cresta Stated: {polygon_stated}<br>Count: {count}<br>Latitude: {lat_long[0]}<br>Longitude: {lat_long[1]}</p>''',
                                    max_width=1500, icon=folium.Icon(color='red')).add_to(outside_cluster)

                    # Add the crestas on to the map as individual zones and as a whole
                    all_crestas = folium.FeatureGroup(name='All Crestas', show=True)
                    for idx, row in self.gdf_filtered.iterrows():
                        folium.GeoJson(row['geometry'], style_function=lambda x: {'fillColor': 'gray', 'color': 'black', 'weight': 1}).add_to(all_crestas)
                    all_crestas.add_to(m)   
                    for idx, row in self.gdf_filtered.iterrows():
                        polygon_group = folium.FeatureGroup(name=row['ZONE_NUM'], show=False)
                        folium.GeoJson(row['geometry'], style_function=lambda x: {'fillColor': 'gray', 'color': 'black', 'weight': 1}).add_to(polygon_group)
                        polygon_group.add_to(m)

                    # add a legend 
                    legend_html = """
                            <div style="position: fixed; bottom:30px; left:30px; z-index:9999; font-size:11px;">
                                <p><i class="fa fa-circle fa-1x" style="color:orange"></i> Cresta/Coordinate Inconsistency </p>
                                <p><i class="fa fa-circle fa-1x" style="color:red"></i> Outside All Cresta Zones </p>
                            </div>
                            """
                    
                    # Adding various tiles option incorporated within Folium library
                    folium.TileLayer('CartoDB positron').add_to(m) # Selected as Default
                    folium.TileLayer('OpenStreetMap').add_to(m) 

                    # Layer control for each layers added can be checked & unchecked while selecting
                    folium.LayerControl(position='topright',collapsed=True, autoZIndex=True).add_to(m)

                    m.get_root().html.add_child(folium.Element(legend_html))

                    # Full Screen option
                    folium.plugins.Fullscreen().add_to(m)

                    m.save(file_save_path)

                    self.append_text('Map Output Successfully Saved')
                    self.update_progress(100)
    
        except Exception:

            self.append_text('Map Output Could Not be Made')
            self.update_progress(100)

    def reverse_geocode(self):

        self.update_progress(1)

        if self.analysis_just_run == True:
            if len(self.outside_df_country_check) >= 1:
                df = self.outside_df_country_check
                geolocator = Nominatim(user_agent="geoapiExercises")
                locations = []
                already_checked = {}
                for idx, row in df.iterrows():
                    try:
                        location = geolocator.reverse(f"{row['LAT_AJG']}, {row['LONG_AJG']}")
                        lat_long = (row['LAT_AJG'], row['LONG_AJG'])
                        if lat_long in already_checked:
                            value = already_checked[lat_long]
                            locations.append(value)
                            continue
                        else:
                            if location:
                                locations.append(location.address)
                            else:
                                locations.append("N/A")
                    except:
                        locations.append("N/A")
                df['Reverse Searched'] = locations
                self.outside_df_country_reversed_checked = df
                
                dir_path = self.excel_path.replace(self.excel_path.split("/")[-1], "")
                outside_df_path = dir_path + '/Incorrect Lat Longs.csv'
                self.outside_df_country_check.to_csv(outside_df_path, index=False)
                self.append_text('Reverse Geocoding Ran Successfully')
                self.update_progress(100)
            else:
                self.append_text('There were no invalid LL to reverse geocode')
                self.update_progress(100)
        else:
            self.append_text('Please run full country analysis first')
            self.update_progress(100)
    
    def clean_data(self):
        
        ''' Standardize Column Names '''

        if self.analysis_type == 'Cresta':
            
            lat_col = self.lat_col
            long_col = self.long_col
            cresta_col = self.cresta_col
            locnum_col = self.loc_col
            self.data = self.data.rename(columns={lat_col: 'LAT_AJG', long_col: 'LONG_AJG', cresta_col: 'CRESTA_AJG', locnum_col: 'LOCNUM_AJG'})

        else:
            
            lat_col = self.lat_col
            long_col = self.long_col
            locnum_col = self.loc_col
            self.data = self.data.rename(columns={lat_col: 'LAT_AJG', long_col: 'LONG_AJG', locnum_col: 'LOCNUM_AJG'})
        
        col_names = self.data.columns.tolist()
        for i in col_names:
            col_name = i.split(':')[0]
            if col_name == 'Unnamed':
                self.data = self.data.drop(i, axis=1)
       
        column_names = None
        if self.analysis_type == 'Cresta':
            column_names = ['LAT_AJG', 'LONG_AJG', 'CRESTA_AJG', 'LOCNUM_AJG']
        else:
            column_names = ['LAT_AJG', 'LONG_AJG', 'LOCNUM_AJG']

        if self.analysis_just_run == False: 
            self.nan_data = pd.DataFrame(columns=self.data.columns)
        elif self.cresta_just_run == 1:
            if self.analysis_type == 'Country':
                self.nan_data = pd.DataFrame(columns=self.data.columns)

        first_1000_rows_skip = []

        counter = 0
        total = len(self.data)
        if self.analysis_type == 'Country First 1000 Rows':
            total = 1000
        
        for idx, row in self.data.iterrows():

            counter += 1
            print(f"Data Cleaned: {counter}/{total}")

            if self.analysis_type == 'Country First 1000 Rows':
                if counter == 1000:
                    break
            
            ''' Check for Nulls '''
            
            # remove rows which dont have values in important columns
            if row[column_names].isnull().any():
                if self.analysis_type == 'Country First 1000 Rows':
                    if counter <= 1000:
                        first_1000_rows_skip.append(idx)
                else:
                    self.nan_data = pd.concat([self.nan_data, row.to_frame().T])
                    self.data.drop(idx, inplace=True)
                    continue

            if type(row['LOCNUM_AJG']) == str:
                if row['LOCNUM_AJG'][0] == ' ':
                    if self.analysis_type == 'Country First 1000 Rows':
                        if counter <= 1000:
                            first_1000_rows_skip.append(idx)
                    else:
                        self.nan_data = pd.concat([self.nan_data, row.to_frame().T])
                        self.data.drop(idx, inplace=True)
                        continue
                
            ''' Check for Wrong Data Type '''
            try:
                self.data.at[idx, 'LAT_AJG'] = float(row['LAT_AJG'])
                self.data.at[idx, 'LONG_AJG'] = float(row['LONG_AJG'])
            except Exception:
                if self.analysis_type == 'Country First 1000 Rows':
                    if counter <= 1000:
                        first_1000_rows_skip.append(idx)
                else:
                    self.nan_data = pd.concat([self.nan_data, row.to_frame().T])
                    self.data.drop(idx, inplace=True)
                    continue
                
            if self.analysis_type == 'Cresta':
                num = row['CRESTA_AJG']
                try:
                    num = int(num)
                    self.data.at[idx, 'CRESTA_AJG'] = num
                except Exception:
                    try:
                        num = num.split('.')[0]
                        try:
                            num = int(num)
                            self.data.at[idx, 'CRESTA_AJG'] = num
                        except Exception:
                            if self.analysis_type == 'Country First 1000 Rows':
                                if counter <= 1000:
                                    first_1000_rows_skip.append(idx)
                            else:
                                self.nan_data = pd.concat([self.nan_data, row.to_frame().T])
                                self.data.drop(idx, inplace=True)
                                continue
                            
                    except Exception:
                        if self.analysis_type == 'Country First 1000 Rows':
                            if counter <= 1000:
                                first_1000_rows_skip.append(idx)
                        else:
                            self.nan_data = pd.concat([self.nan_data, row.to_frame().T])
                            self.data.drop(idx, inplace=True)
                            continue

        self.first_1000_idx = first_1000_rows_skip
                    
        ''' Map Cresta for Dominican Republic '''
                    
        if self.analysis_type == 'Cresta':
            if self.country_name == 'Dominican Republic':
                value_mapping = {'a': 1, 'b': 2, 'c': 3, 'd': 4, 'e': 5}
                for idx, row in self.data.iterrows():
                    if row['CRESTA_AJG'] in value_mapping:
                        self.data.at[idx, 'CRESTA_AJG'] = value_mapping[row['CRESTA_AJG']]
                
    def save_dataframe(self, df, save_path, myformats=[]):

        print('saving dataframe')

        start_time = time.time()
        if len(df.columns) <= 0:
            return
        Nd = len(df.columns)
        Nd_1 = Nd - 1
        formats = myformats[:] # take a copy to modify it
        Nf = len(formats)
        # make sure we have formats for all columns
        if Nf < Nd:
            for ii in range(Nf,Nd):
                coltype = df[df.columns[ii]].dtype
                ff = '%s'
                if coltype == np.int64:
                    ff = '%d'
                elif coltype == np.float64:
                    ff = '%f'
                formats.append(ff)
        with open(save_path, 'w', encoding='latin-1') as fh:
            fh.write(','.join(df.columns) + '\n')
            for row in df.itertuples(index=False):
                ss = ''
                for ii in range(Nd):
                    ss += formats[ii] % row[ii]
                    if ii < Nd_1:
                        ss += ','
                fh.write(ss+'\n')
        end_time = time.time()
        elapsed_time = end_time-start_time
        print(f"Time taken to download file: {elapsed_time:.1f}")

    def geospatial_analysis(self):

        try:

            start_time = time.time()
            
            self.append_text('')
            self.append_text('Geospatial Analysis Started')

            if self.cresta_just_run == 1:
                if self.excel_path and self.excel_path.split(".")[-1] in ["xlsx", "xls"]:
                    self.data = pd.read_excel(self.excel_path)
                elif self.excel_path and self.excel_path.split(".")[-1] in ["csv"]:
                    self.data = pd.read_csv(self.excel_path, encoding='latin-1')
            
            if self.analysis_type != 'Cresta':

                ''' Clean the Data '''
                self.clean_data()
                self.append_text('Data Standardized and Cleaned')
                self.update_progress(10)

                self.cresta_just_run = None
                self.analysis_just_run = False
                self.first_country_name = self.country_name
            
                ''' Initialise Variables '''
                inside_id = []
                outside_low_res = pd.DataFrame(columns=self.data.columns)
                outside_bounds = pd.DataFrame(columns=self.data.columns)
                inside_count = {}
                outside_count = {}
                
                ''' Are the Coordinates within the Bounds? '''
                minx, maxx, miny, maxy = self.gdf_low_res.iloc[0].geometry.bounds
                x_list = [minx, maxx]
                x_list.sort()
                y_list = [miny, maxy]
                y_list.sort() 

                counter = 0

                centroid_counter = 0
                centroid = self.country_geometry_low_res.centroid
                zero_zero_counter = 0
                total = len(self.data)
                
                for idx, row in self.data.iterrows():

                    if self.analysis_type == 'Country First 1000 Rows':
                        if counter == 1000:
                            break
                        else:
                            counter += 1
                            print(f"Low Resolution: {counter}/{1000}")
                        if idx in self.first_1000_idx:
                            continue
                    else:
                        counter += 1
                        print(f"Low Resolution: {counter}/{total}")
                    
                    location = shapely.geometry.Point(row['LONG_AJG'], row['LAT_AJG'])
                    lat_long = (row['LAT_AJG'], row['LONG_AJG'])

                    if location == centroid:
                        centroid_counter += 1

                    if lat_long == (0,0):
                        zero_zero_counter += 1

                    if lat_long in inside_count:
                        inside_count[lat_long] = inside_count.get(lat_long, 0) + 1
                        inside_id.append(row['LOCNUM_AJG'])
                        continue
                    elif lat_long in outside_count:
                        outside_count[lat_long] = outside_count.get(lat_long, 0) + 1
                        outside_bounds = pd.concat([outside_bounds, row.to_frame().T])
                        continue

                    # quickly remove data in incompatible zones
                    if x_list[0] <= row['LONG_AJG'] <= x_list[1] == False:
                        outside_bounds = pd.concat([outside_bounds, row.to_frame().T])
                        continue
                    if y_list[0] <= row['LAT_AJG'] <= y_list[1] == False:
                        outside_bounds = pd.concat([outside_bounds, row.to_frame().T])
                        continue
                    
                    if location.within(self.country_geometry_buffer):
                        inside_count[lat_long] = inside_count.get(lat_long, 0) + 1
                        inside_id.append(row['LOCNUM_AJG'])
                    else:
                        outside_low_res = pd.concat([outside_low_res,  row.to_frame().T], ignore_index=True)
                    
                self.update_progress(60)
                
                counter_2 = 0
                total = len(outside_low_res)
                
                for idx, row in outside_low_res.iterrows():

                    counter_2 += 1
                    print(f"High Resolution: {counter_2}/{total}")
            
                    location = shapely.geometry.Point(row['LONG_AJG'], row['LAT_AJG'])
                    lat_long = (row['LAT_AJG'], row['LONG_AJG'])
                            
                    # add to our variables
                    if location.within(self.country_geometry_high_res):
                        inside_count[lat_long] = inside_count.get(lat_long, 0) + 1
                        inside_id.append(row['LOCNUM_AJG'])
                        outside_low_res.drop(idx, inplace=True)
                    else:
                        outside_count[lat_long] = outside_count.get(lat_long, 0) + 1
                    
                inside_df = self.data[self.data['LOCNUM_AJG'].isin(inside_id)]
                outside_low_res = pd.concat([outside_low_res, outside_bounds])
                
                self.inside_df_country_check = inside_df
                self.outside_df_country_check = outside_low_res[['LOCNUM_AJG', 'LAT_AJG', 'LONG_AJG']]
                self.country_inside_dict = inside_count
                self.country_outside_dict = outside_count

                total = len(self.inside_df_country_check) + len(self.outside_df_country_check) + len(self.nan_data)
                if self.analysis_type == 'Country First 1000 Rows':
                    if len(self.data) <= 1000:
                        total = len(self.data)
                    else:
                        total = 1000
                total_inside = len(self.inside_df_country_check)
                total_outside = len(self.outside_df_country_check)
                total_nan = None
                if self.analysis_type == 'Country':
                    total_nan = len(self.nan_data)
                else:
                    total_nan = len(self.first_1000_idx)
                
                self.append_text('')
                self.append_text('Summary Review')
                self.append_text(f'Total Assessed: {total}')
                self.append_text(f'Total Inside: {total_inside}, {((total_inside)*100/total):.2f}%')
                self.append_text(f'Total Outside: {total_outside}, {((total_outside)*100/total):.2f}%')
                self.append_text(f'Total Data Type Error/Null: {total_nan}, {((total_nan)*100/total):.2f}%')
                self.append_text(f'% of Points at Country Centroid: {((centroid_counter*100)/total):.2f}%')
                self.append_text(f'% of Points at (0,0): {((zero_zero_counter*100)/total):.2f}%')
                self.append_text('')

                if self.analysis_type == 'Country':
                    self.analysis_just_run = True
                    self.first_country_name = self.country_name

                inside_df = inside_df[['LOCNUM_AJG', 'LAT_AJG', 'LONG_AJG']]
                nan_to_print = self.nan_data[['LOCNUM_AJG', 'LAT_AJG', 'LONG_AJG']]

                if self.analysis_type == 'Country':

                    dir_path = self.excel_path.replace(self.excel_path.split("/")[-1], "")
                    inside_df_path = dir_path + '/Correct Lat Longs.csv'
                    self.save_dataframe(inside_df, inside_df_path, myformats=['%s','%f', '%f'])
                    outside_df_path = dir_path + '/Incorrect Lat Longs.csv'
                    self.save_dataframe(self.outside_df_country_check, outside_df_path, myformats=['%s','%f', '%f'])
                    nan_data_path = dir_path + '/Null or Wrong Data Type Lat Longs.csv'
                    self.save_dataframe(nan_to_print, nan_data_path, myformats=['%s','%f', '%f'])

                if self.analysis_type == 'Country':
                    self.recent_analysis = 'Country'
                self.update_progress(100)
            
            if self.analysis_type == 'Cresta':
            
                data = None
                before_cleaned = None
                if self.analysis_just_run == True:
                    if self.country_name == self.first_country_name:
                        self.data = self.inside_df_country_check
                        before_cleaned = len(self.data)
                        self.clean_data()
                        data = self.data
    
                else:
                    self.clean_data()
                    data = self.data
                    
                # initialise variables
                inside = {}
                inside_id = []
                cresta_ambiguity = {}
                cresta_ambiguity_id = []
                outside = {}
                outside_id = []

                # create unified polygons
                gdf_filtered = self.gdf[['ZONE_NUM', 'geometry']]
                gdf_filtered['ZONE_NUM'] = gdf_filtered['ZONE_NUM'].astype(int)

                centroid_counter = 0
                centroids = []
                for idx, row in gdf_filtered.iterrows():
                    centroid = row['geometry'].centroid
                    centroids.append(centroid)

                self.gdf_filtered = gdf_filtered

                # initialise a spatial index 
                my_spatial_index = Index()
                for id, country in gdf_filtered.iterrows():
                    my_spatial_index.insert(id, country.geometry.bounds)

                available_crestas = gdf_filtered['ZONE_NUM'].unique().tolist()
                available_crestas_int = list(map(int, available_crestas)) 

                counter_cresta = 0

                for idx, row in data.iterrows():
                    
                    counter_cresta += 1

                    point = shapely.geometry.Point(row['LONG_AJG'], row['LAT_AJG']) 
                    lat_long = (row['LAT_AJG'], row['LONG_AJG'])

                    if point in centroids:
                        centroid_counter += 1

                    client_cresta_num = int(row['CRESTA_AJG'])

                    key = (lat_long, client_cresta_num)
                    if key in inside:
                        inside[key]['count'] += 1
                        inside_id.append(row['LOCNUM_AJG'])
                        continue
                    elif key in outside:
                        outside[key]['count'] += 1
                        outside_id.append(row['LOCNUM_AJG'])
                        continue
                    elif key in cresta_ambiguity:
                        cresta_ambiguity[key]['count'] += 1
                        cresta_ambiguity_id.append(row['LOCNUM_AJG'])
                        continue

                    if client_cresta_num not in available_crestas_int:
                        possible_matches_id = list(my_spatial_index.intersection(point.bounds))
                        if possible_matches_id:
                            for row_id in possible_matches_id:
                                if gdf_filtered.geometry[row_id].contains(point):
                                    key = (lat_long, client_cresta_num)
                                    cresta_ambiguity[key] = cresta_ambiguity.get(key, {'polygon_stated': client_cresta_num, 'polygon_found': gdf_filtered['ZONE_NUM'][row_id], 'count': 0})
                                    cresta_ambiguity[key]['count'] += 1
                                    cresta_ambiguity_id.append(row['LOCNUM_AJG'])
                                else: 
                                    continue      
                        else:
                            key = (lat_long, client_cresta_num)
                            outside[key] = outside.get(key, {'polygon_stated': client_cresta_num, 'count': 0})
                            outside[key]['count'] += 1
                            outside_id.append(row['LOCNUM_AJG'])                    

                    else:

                        cresta_geometry = self.gdf_filtered.loc[self.gdf_filtered['ZONE_NUM'] == client_cresta_num, 'geometry'].values[0]

                        if point.within(cresta_geometry):
                            key = (lat_long, client_cresta_num)
                            inside[key] = inside.get(key, {'polygon': client_cresta_num, 'count': 0})
                            inside[key]['count'] += 1
                            inside_id.append(row['LOCNUM_AJG'])

                        else: 
                            possible_matches_id = list(my_spatial_index.intersection(point.bounds))
                            if possible_matches_id:
                                for row_id in possible_matches_id:
                                    if gdf_filtered.geometry[row_id].contains(point):
                                        key = (lat_long, client_cresta_num)
                                        cresta_ambiguity[key] = cresta_ambiguity.get(key, {'polygon_stated': client_cresta_num, 'polygon_found': gdf_filtered['ZONE_NUM'][row_id], 'count': 0})
                                        cresta_ambiguity[key]['count'] += 1
                                        cresta_ambiguity_id.append(row['LOCNUM_AJG'])
                                    else: 
                                        continue      

                            else:
                                key = (lat_long, client_cresta_num)
                                outside[key] = outside.get(key, {'polygon_stated': client_cresta_num, 'count': 0})
                                outside[key]['count'] += 1
                                outside_id.append(row['LOCNUM_AJG'])

                self.inside_df_cresta_check = data.loc[data['LOCNUM_AJG'].isin(inside_id)]
                self.outside_df_cresta_check = data.loc[data['LOCNUM_AJG'].isin(outside_id)]
                self.ambiguous_df_cresta_check = data.loc[data['LOCNUM_AJG'].isin(cresta_ambiguity_id)]
                self.inside_df_cresta_check = self.inside_df_cresta_check[['LOCNUM_AJG', 'LAT_AJG', 'LONG_AJG', 'CRESTA_AJG']]
                self.outside_df_cresta_check = self.outside_df_cresta_check[['LOCNUM_AJG', 'LAT_AJG', 'LONG_AJG', 'CRESTA_AJG']]
                self.ambiguous_df_cresta_check = self.ambiguous_df_cresta_check[['LOCNUM_AJG', 'LAT_AJG', 'LONG_AJG', 'CRESTA_AJG']]
                self.nan_data = self.nan_data[['LOCNUM_AJG', 'LAT_AJG', 'LONG_AJG', 'CRESTA_AJG']]

                self.cresta_inside_dict = inside
                self.cresta_outside_dict = outside
                self.cresta_ambiguous_dict = cresta_ambiguity

                if len(self.ambiguous_df_cresta_check) > 0:
                    for id, row in self.ambiguous_df_cresta_check.iterrows():
                        lat_long = (row['LAT_AJG'], row['LONG_AJG'])
                        value = cresta_ambiguity[(lat_long, row['CRESTA_AJG'])]['polygon_found']
                        self.ambiguous_df_cresta_check.at[id, 'CRESTA_FOUND_AJG'] = value.item() if isinstance(value, numpy.int64) else value
                
                total_inside = len(self.inside_df_cresta_check)
                total_outside = len(self.outside_df_cresta_check)
                total_ambiguous = len(self.ambiguous_df_cresta_check)
                total_null = len(self.nan_data)
                if self.analysis_just_run == True:
                    total_null = before_cleaned - len(data)
                total = len(data) + total_null

                dir_path = self.excel_path.replace(self.excel_path.split("/")[-1], "")
                inside_df_path = dir_path + '/Consistent Cresta Lat Longs.csv'
                self.save_dataframe(self.inside_df_cresta_check, inside_df_path, myformats=['%s','%f', '%f', '%s'])
                outside_df_path = dir_path + '/Incorrect cresta Lat Longs.csv'
                self.save_dataframe(self.outside_df_cresta_check, outside_df_path, myformats=['%s','%f', '%f', '%s'])
                ambiguous_df_path = dir_path + '/Inconsistent Cresta Lat Longs.csv'
                self.save_dataframe(self.ambiguous_df_cresta_check, ambiguous_df_path, myformats=['%s','%f', '%f'])
                nan_data_path = dir_path + '/Null or Wrong Data Type Cresta Lat Longs.csv'
                self.save_dataframe(self.nan_data, nan_data_path, myformats=['%s','%f', '%f', '%s'])

                self.append_text('')
                self.append_text('Cresta Summary Review')
                self.append_text(f'Total Assessed: {total}')
                if self.analysis_just_run == True:
                    self.append_text(f'Total Excluded from Analysis as Outside of Country/Null in Previous Analysis: {len(self.outside_df_country_check) + (len(self.nan_data) - total_null)}')
                self.append_text(f'Total Inside: {total_inside}, {((total_inside)*100/total):.2f}%')
                self.append_text(f'Total Outside: {total_outside}, {((total_outside)*100/total):.2f}%')
                self.append_text(f'Total Ambiguous: {total_ambiguous}, {((total_ambiguous)*100/total):.2f}%')
                if self.analysis_just_run == True:
                    self.append_text(f'Additional Data Type Error/Null Found: {total_null}, {((total_null)*100/total):.2f}%')
                else:
                    self.append_text(f'Total Data Type Error/Null: {total_null},  {((total_null)*100/total):.2f}%')
                self.append_text(f'% of Points at Cresta Centroid: {((centroid_counter*100)/total):.2f}%')
                self.append_text('')

                self.cresta_just_run = 1
                self.recent_analysis = 'Cresta'
                
                self.update_progress(100)
        
            if self.analysis_type != 'Country First 1000 Rows':
                self.map_output_button.setEnabled(True)      
        
            end_time = time.time()
            elapsed_time = end_time - start_time
            self.append_text('')
            self.append_text(f'Elapsed time: {(elapsed_time/60):.1f} mins')
            self.append_text('')
            
        except Exception as e:

            self.update_progress(100)

            my_list = None
            if self.analysis_type == 'Cresta':
                my_list = [self.data, self.cresta_col, self.loc_col, self.lat_col, self.long_col]
            else:
                my_list = [self.data, self.loc_col, self.lat_col, self.long_col]
            for i in my_list:
                if i is None:
                    self.append_text(f'Not all fields were satisfied. Please ensure all inputs necessary have been inputted correctly')
                    break
                else:
                    continue
            else:
                self.append_text(f'Error during geospatial analysis: {str(e)}')

    def get_cresta_shp(self):

        country_with_underscore = self.country_name.replace(' ', '_')

        try:
            if os.path.exists(r'J:\cms\Internal\Territories\LAC\3. Vendor & Internal Models\LAC Geodata Validation Tools\LAC CRESTA ZONES PER COUNTRY - DO NOT MOVE\South American Cresta Zones Per Country High Res'):
                shp_path = rf'J:\cms\Internal\Territories\LAC\3. Vendor & Internal Models\LAC Geodata Validation Tools\LAC CRESTA ZONES PER COUNTRY - DO NOT MOVE\South American Cresta Zones Per Country High Res\{self.country_name}\{country_with_underscore}.shp'
                self.append_text(f'.shp download started')
                self.gdf = gpd.read_file(shp_path)
                self.append_text(f'Country successfully downloaded')
                self.update_progress(10)
            elif os.path.exists(r'J:\Analytics\cms\Internal\Territories\LAC\3. Vendor & Internal Models\LAC Geodata Validation Tools\LAC CRESTA ZONES PER COUNTRY - DO NOT MOVE\South American Cresta Zones Per Country High Res'):
                shp_path = rf'J:\Analytics\cms\Internal\Territories\LAC\3. Vendor & Internal Models\LAC Geodata Validation Tools\LAC CRESTA ZONES PER COUNTRY - DO NOT MOVE\South American Cresta Zones Per Country High Res\{self.country_name}\{country_with_underscore}.shp'
                self.append_text(f'.shp download started')
                self.gdf = gpd.read_file(shp_path)
                self.append_text(f'Country successfully downloaded')
                self.update_progress(10)

        except Exception:
            self.append_text('The file location appears to have moved. Ensure the following path file still exists "J:\cms\Internal\Territories\LAC\3. Vendor & Internal Models\LAC Geodata Validation Tools\LAC CRESTA ZONES PER COUNTRY - DO NOT MOVE\South American Cresta Zones Per Country High Res". Note, your file path may start with "Analytics"')
            self.update_progress(100)
            
    def get_country_shp(self):
        
        try:

            if self.country_name:
                
                if os.path.exists(r'J:\cms\Internal\Territories\LAC\3. Vendor & Internal Models\LAC Geodata Validation Tools'):
                    base_path = r'J:\cms\Internal\Territories\LAC\3. Vendor & Internal Models\LAC Geodata Validation Tools'
                    if self.country_name == 'Saint Vincent and the Grenadines':
                        self.shapefile_low_res = os.path.join(base_path, 'LAC COUNTRY SHAPEFILES - DO NOT MOVE',
                                                        '3. World MidRes1', 'Saint Vincent and the Grenadines',
                                                        'Saint Vincent and the Grenadines.shp')
                    else:
                        self.shapefile_low_res = os.path.join(base_path, 'LAC COUNTRY SHAPEFILES - DO NOT MOVE',
                                                        '2. World LowRes', self.country_name, f'{self.country_name}.shp')
                    self.shapefile_high_res = os.path.join(base_path, 'LAC COUNTRY SHAPEFILES - DO NOT MOVE',
                                                        '5. World HighRes',self.country_name, f'{self.country_name}.shp')
                
                elif os.path.exists(r'J:\Analytics\cms\Internal\Territories\LAC\3. Vendor & Internal Models\LAC Geodata Validation Tools'):
                    base_path = r'J:\Analytics\cms\Internal\Territories\LAC\3. Vendor & Internal Models\LAC Geodata Validation Tools'
                    if self.country_name == 'Saint Vincent and the Grenadines':
                        self.shapefile_low_res = os.path.join(base_path, 'LAC COUNTRY SHAPEFILES - DO NOT MOVE',
                                                        '3. World MidRes1', 'Saint Vincent and the Grenadines',
                                                        'Saint Vincent and the Grenadines.shp')
                    else:
                        self.shapefile_low_res = os.path.join(base_path, 'LAC COUNTRY SHAPEFILES - DO NOT MOVE',
                                                        '2. World LowRes', self.country_name, f'{self.country_name}.shp')
                    self.shapefile_high_res = os.path.join(base_path, 'LAC COUNTRY SHAPEFILES - DO NOT MOVE',
                                                        '5. World HighRes', self.country_name, f'{self.country_name}.shp')
                else:
                    self.append_Text('File path to shapefiles cannot be found')
                    
                # get all the geometries and shapefiles
                self.gdf_low_res = gpd.read_file(self.shapefile_low_res)
                self.gdf_high_res = gpd.read_file(self.shapefile_high_res) 
                self.country_gdf_low_res = self.gdf_low_res.loc[self.gdf_low_res['NAME'] == self.country_name]
                self.country_geometry_low_res = self.country_gdf_low_res.geometry.iloc[0]
                self.country_geometry_buffer = self.country_geometry_low_res.buffer(-0.01)
                self.country_gdf_high_res = self.gdf_high_res.loc[self.gdf_high_res['NAME'] == self.country_name]
                self.country_geometry_high_res = self.country_gdf_high_res.geometry.iloc[0]
                        
            else:
                self.append_text('No Country Has Been Selected')

        except Exception:

            self.append_text('The file location appears to have moved. Ensure the following path file still exists "J:\cms\Internal\Territories\LAC\3. Vendor & Internal Models\LAC Geodata Validation Tools\LAC COUNTRY SHAPEFILES - DO NOT MOVE\". Note, your file path may start with "Analytics"')
            self.update_progress(100)
            
    def update_progress(self, progress_value):
        self.progress_bar.setValue(progress_value)    
        
    def standardize_cols(self):
        
        self.loc_col = self.listbox3.currentText()
        self.lat_col = self.listbox4.currentText()
        self.long_col = self.listbox5.currentText()
        self.cresta_col = self.listbox6.currentText()
        self.append_text(f'Locnum Column selected: {self.loc_col}')
        self.append_text(f'Latitude Column selected: {self.lat_col}')
        self.append_text(f'Longitude Column selected: {self.long_col}')
        if self.analysis_type == 'Cresta':
            self.append_text(f'Cresta Column selected: {self.cresta_col}')

        if self.analysis_just_run == False:
            if self.analysis_type != 'Cresta':
                self.data = self.data_copy

    def update_listbox(self):
        
        self.listbox3.clear()
        self.listbox4.clear()
        self.listbox5.clear()
        self.listbox6.clear()
        columns = self.data.columns.tolist()
        self.listbox3.addItems(columns)
        self.listbox4.addItems(columns)
        self.listbox5.addItems(columns)
        self.listbox6.addItems(columns)
    
    def choose_country(self):
        
        selected_country = self.listbox2.currentText()
        if selected_country:
            self.country_name = selected_country
            self.append_text(f'Country Selected: {self.country_name}')
            self.update_progress(2)
        else:
            self.append_text('Please select a country.')

        if self.analysis_just_run == True:
            if self.country_name != self.first_country_name:
                self.append_text('Warning: The country selected does not match the previously used country')

        ''' Download shapefiles '''
        if self.analysis_type != 'Cresta':
            self.get_country_shp()
            self.append_text('.shp Downloaded')
            self.update_progress(5)     
        else:
            self.get_cresta_shp()
            self.append_text('.shp Downloaded')
        
    def update_grey(self, selected_output):
        
        if selected_output:
            
            self.get_available_countries()
            
            if self.analysis_type == 'Cresta':
                
                self.label7.setEnabled(True)
                self.listbox6.setEnabled(True)
                self.reverse_geocode_btn.setEnabled(False) 
                
            else:
                
                self.label7.setEnabled(False)
                self.listbox6.setEnabled(False)
                self.reverse_geocode_btn.setEnabled(True) 
                
        else:
            self.append_text('Please select an analysis type.')
    
    def choose_analysis(self):
       
        selected_output = self.listbox.currentText()
        self.analysis_type = selected_output
        self.append_text(f'Analysis type selected: {self.analysis_type}')
        self.update_grey(self.analysis_type)
        self.update_progress(1)
        
    def get_available_countries(self):
        
        if self.analysis_type:
            
            if self.analysis_type in ['Country First 1000 Rows', 'Country']:
            
                # get the path_file for country level .shp
                countries_1 = None
                if os.path.exists(r'J:\cms\Internal\Territories\LAC\3. Vendor & Internal Models\LAC Geodata Validation Tools\LAC COUNTRY SHAPEFILES - DO NOT MOVE'):
                    countries_1 = glob.glob(r'J:\cms\Internal\Territories\LAC\3. Vendor & Internal Models\LAC Geodata Validation Tools\LAC COUNTRY SHAPEFILES - DO NOT MOVE\2. World LowRes\*\\')
                elif os.path.exists(r'J:\Analytics\cms\Internal\Territories\LAC\3. Vendor & Internal Models\LAC Geodata Validation Tools\LAC COUNTRY SHAPEFILES - DO NOT MOVE'):
                    countries_1 = glob.glob(r'J:\Analytics\cms\Internal\Territories\LAC\3. Vendor & Internal Models\LAC Geodata Validation Tools\LAC COUNTRY SHAPEFILES - DO NOT MOVE\2. World LowRes\*\\')
                countries_list_1 = []
                for string in countries_1:
                    split_parts = string.rsplit("\\", 2)
                    country = split_parts[1]
                    countries_list_1.append(country)
                
                self.listbox2.clear()
                self.listbox2.addItems(countries_list_1)
                 
            if self.analysis_type == 'Cresta':
            
                # get the path_file for the cresta level .shp
                countries_2 = None
                if os.path.exists(r'J:\cms\Internal\Territories\LAC\3. Vendor & Internal Models\LAC Geodata Validation Tools\LAC CRESTA ZONES PER COUNTRY - DO NOT MOVE\South American Cresta Zones Per Country High Res'):
                    countries_2 = glob.glob(r'J:\cms\Internal\Territories\LAC\3. Vendor & Internal Models\LAC Geodata Validation Tools\LAC CRESTA ZONES PER COUNTRY - DO NOT MOVE\South American Cresta Zones Per Country High Res\*\\')
                elif os.path.exists(r'J:\Analytics\cms\Internal\Territories\LAC\3. Vendor & Internal Models\LAC Geodata Validation Tools\LAC CRESTA ZONES PER COUNTRY - DO NOT MOVE\South American Cresta Zones Per Country High Res'):
                    countries_2 = glob.glob(r'J:\Analytics\cms\Internal\Territories\LAC\3. Vendor & Internal Models\LAC Geodata Validation Tools\LAC CRESTA ZONES PER COUNTRY - DO NOT MOVE\South American Cresta Zones Per Country High Res\*\\')
                countries_list_2 = []
                for string in countries_2:
                    split_parts = string.rsplit("\\", 2)
                    country = split_parts[1]
                    countries_list_2.append(country)
            
                self.listbox2.clear()
                self.listbox2.addItems(countries_list_2)
             
    def get_excel_doc(self):

        self.update_progress(1)
        
        self.country_name = None
        self.first_country_name = None
        self.analysis_type = None
        self.excel_path = None
        self.data = None
        self.loc_col = None
        self.lat_col = None
        self.long_col = None
        self.cresta_col = None
        
        self.nan_data = None
        self.inside_df_country_check = None
        self.outside_df_country_check = None
        self.outside_df_country_reversed_checked = None
        self.inside_df_cresta_check = None
        self.outside_df_cresta_check = None
        self.ambiguous_df_cresta_check = None
        self.country_inside_dict = None
        self.country_outside_dict = None
        self.cresta_inside_dict = None
        self.cresta_outside_dict = None
        self.cresta_ambiguous_dict = None
        
        self.first_1000_idx = None
        self.cresta_just_run = None
        self.analysis_just_run = False
        self.recent_analysis = None

        self.excel_path, _ = QFileDialog.getOpenFileName(self, 'Select Excel File')
        if self.excel_path and self.excel_path.split(".")[-1] in ["xlsx", "xls"]:
            self.data = pd.read_excel(self.excel_path)
            self.data_copy = self.data
            self.append_text('Excel file selected.')
            self.update_listbox()
        elif self.excel_path and self.excel_path.split(".")[-1] in ["csv"]:
            self.data = pd.read_csv(self.excel_path, encoding='latin-1')
            self.data_copy = self.data
            self.append_text('Excel file selected.')
            self.update_listbox()
        else:
            self.append_text('Please choose a valid Excel file.')
            
    def append_text(self, text):
        self.textOutput.append(text)
        self.textOutput.verticalScrollBar().setValue(self.textOutput.verticalScrollBar().maximum())
        self.textOutput.update()
        
def main():
    app = QApplication(sys.argv)
    window = CoordinateValidationTool()
    return app.exec_()

if __name__ == '__main__':
    
    exit_code = main()
    sys.exit(exit_code)  
