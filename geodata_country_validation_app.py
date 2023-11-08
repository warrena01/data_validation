import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QLabel, QFileDialog, QComboBox, QTextEdit, QProgressBar
from PyQt5.QtCore import Qt
import geopandas as gpd
import pandas as pd
import folium
from shapely.geometry import Point
from folium.plugins import HeatMap
from os import path
from geopy.geocoders import Nominatim
import glob
import openpyxl
import xlrd
import time 

class CoordinateValidationTool(QMainWindow):

    def __init__(self):
        super().__init__()

        self.excel_path = None
        self.data = None
        self.shapefile_low_res = None
        self.shapefile_high_res = None
        self.country_name = None
        self.file_path = None
        self.lat_col = None
        self.long_col = None
        self.loc_col = None
        
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setGeometry(400, 410, 250, 20)
        self.progress_bar.setOrientation(Qt.Horizontal)
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0) 

        self.initUI()

    def initUI(self):
        self.setWindowTitle('Coordinate Validation Tool')
        self.setGeometry(250, 250, 650, 650)

        '''GET EXCEL'''
        self.label1 = QLabel('Step 1: Select the excel workbook', self)
        self.label1.setGeometry(30, 20, 400, 20)
        self.btn1 = QPushButton('Choose File', self)
        self.btn1.setGeometry(30, 50, 100, 30)
        self.btn1.clicked.connect(self.get_excel_doc)

        '''GET COUNTRY'''
        self.label2 = QLabel('Step 2: Select the country, then hit confirm', self)
        self.label2.setGeometry(30, 100, 400, 20)
        self.listbox = QComboBox(self)
        self.listbox.setGeometry(30, 130, 200, 30)
        countries = glob.glob(r'J:\cms\External\Team Members\Adam Warren\LAC COUNTRY SHAPEFILES - DO NOT MOVE\2. World LowRes\*\\')
        countries_list = []
        for string in countries:
            split_parts = string.rsplit("\\", 2)
            country = split_parts[1]
            countries_list.append(country)
        self.listbox.addItems(countries_list)
        self.listbox.setCurrentIndex(-1)
        self.btn2 = QPushButton('Confirm', self)
        self.btn2.setGeometry(30, 160, 100, 30)
        self.btn2.clicked.connect(self.choose_country)
        
        '''GET OUTPUT MAP'''
        self.label3 = QLabel('Step 3: Select the output map, then hit confirm', self)
        self.label3.setGeometry(30, 210, 500, 20)
        self.listbox2 = QComboBox(self)
        self.listbox2.setGeometry(30, 240, 200, 30)
        self.listbox2.addItems(['Heatmap', 'Heatmap Overlay', 'Point Data Map', 'All'])
        self.listbox2.setCurrentIndex(-1)
        self.btn3 = QPushButton('Confirm', self)
        self.btn3.setGeometry(30, 270, 100, 30)
        self.btn3.clicked.connect(self.choose_output)

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

        self.btn9 = QPushButton('Confirm', self)
        self.btn9.setGeometry(150, 440, 100, 30)
        self.btn9.clicked.connect(self.standardize_cols)

        self.label8 = QLabel('Step 5: Start Analysis', self)
        self.label8.setGeometry(20, 490, 300, 20)

        self.btn5 = QPushButton('Go', self)
        self.btn5.setGeometry(20, 520, 100, 30)
        self.btn5.clicked.connect(self.geospatial_analysis)

        self.textOutput = QTextEdit(self)
        self.textOutput.setGeometry(400, 30, 230, 380)
        self.textOutput.setReadOnly(True)
        
        self.exit_btn = QPushButton('Exit', self)
        self.exit_btn.setGeometry(20, 550, 100, 30)  # Position the button at the top-right corner
        self.exit_btn.clicked.connect(self.exit_app)
        
        self.reverse_geocode_btn = QPushButton('Run Reverse Geocode', self)
        self.reverse_geocode_btn.setGeometry(20, 580, 200, 30)
        self.reverse_geocode_btn.clicked.connect(self.reverse_geocode)

        self.show()

    def get_excel_doc(self):
        self.excel_path, _ = QFileDialog.getOpenFileName(self, 'Select Excel File')
        if self.excel_path and self.excel_path.split(".")[-1] in ["xlsx", "xls"]:
            self.data = pd.read_excel(self.excel_path)
            self.append_text('Excel file selected.')
            file_x = self.excel_path.rsplit('/', 1)[0]
            self.file_path = f"{file_x}/map_output.html"
            self.update_listbox()
        elif self.excel_path and self.excel_path.split(".")[-1] in ["csv"]:
            self.data = pd.read_csv(self.excel_path, encoding='latin-1')
            self.append_text('Excel file selected.')
            file_x = self.excel_path.rsplit('/', 1)[0]
            self.file_path = f"{file_x}/map_output.html"
            self.update_listbox()
        else:
            self.append_text('Please choose a valid Excel file.')

    def choose_country(self):
        selected_country = self.listbox.currentText()
        if selected_country:
            self.country_name = selected_country
            base_path = rf'J:\cms\External\Team Members\Adam Warren'
            if self.country_name == 'Saint Vincent and the Grenadines':
                self.shapefile_low_res = path.join(base_path, 'LAC COUNTRY SHAPEFILES - DO NOT MOVE',
                                                '2. World LowRes', 'Saint Vincent and the Grenadines',
                                                'Saint Vincent and the Grenadines.shp')
                
            else:
                self.shapefile_low_res = path.join(base_path, 'LAC COUNTRY SHAPEFILES - DO NOT MOVE',
                                                '2. World LowRes', self.country_name, f'{self.country_name}.shp')
            self.shapefile_high_res = path.join(base_path, 'LAC COUNTRY SHAPEFILES - DO NOT MOVE',
                                                '5. World HighRes', self.country_name, f'{self.country_name}.shp')
            self.append_text(f'Country selected: {self.country_name}')
        else:
            self.append_text('Please select a country.')

    def choose_output(self):
        selected_output = self.listbox2.currentText()
        if selected_output:
            self.map_output = selected_output
            self.append_text(f'Output map selected: {self.map_output}')
        else:
            self.append_text('Please select an output map type.')

    def append_text(self, text):
        self.textOutput.append(text)
        self.textOutput.verticalScrollBar().setValue(self.textOutput.verticalScrollBar().maximum())
        self.textOutput.update()
        
    def plot_maps(self, data, country_geometry_high_res, map_output, outside_high_res, inside_count, file_path, country_geometry_low_res, inside_df, outside_low_res):    
    
        if len(inside_df) > 0 or len(outside_low_res) > 0:

            if len(inside_df) > 0:
                map_plot = folium.Map(location=[inside_df['latitude'].mean(), inside_df['longitude'].mean()], zoom_starts=4, tiles='CartoDB positron')
            else:
                map_plot = folium.Map(location=[self.data['latitude'].mean(), self.data['longitude'].mean()], zoom_starts=4, tiles='CartoDB positron')

            map_plot = folium.Map(location=[data['latitude'].mean(), data['longitude'].mean()], zoom_starts=4, tiles='CartoDB positron')
            country_geojson = country_geometry_high_res.__geo_interface__
            folium.GeoJson(country_geojson, name='Country Boundary',style_function=lambda feature:{
                                'fillColor': 'lightyellow',
                                'color': 'black',
                                'weight': 1,
                                'fillOpacity': 0.3,      
                                }
                            ).add_to(map_plot)

            heat_data = [[row['latitude'], row['longitude']] for index, row in inside_df.iterrows()]
            # heat_data = [[row['latitude'], row['longitude']] for index, row in data.iterrows() if Point(row['longitude'], row['latitude']).within(country_geometry_low_res)]
            marker_cluster = folium.plugins.MarkerCluster(name="Coordinate Points", options={'disableClusteringAtZoom':10},show=False).add_to(map_plot)

        # overlay group
        overlay_group = folium.FeatureGroup(name='Overlay Layers',show=False)
        overlay_group.add_to(map_plot)

        if map_output == 'Heatmap':
            HeatMap(heat_data,name="HeatMap", show = False).add_to(map_plot)

        elif map_output == 'Heatmap Overlay':
            HeatMap(heat_data,name="HeatMap", show = False).add_to(map_plot)
            for lat_long,count in inside_count.items():
                folium.Circle(lat_long,popup={lat_long},
                                radius=count*5, color ='black', fill=True,fill_opacity=0.5).add_to(overlay_group)

        elif map_output == 'Point Data Map':
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
            folium.Marker(location=[lat,lon], popup=f"Outside Boundaries<br>ID:{ID}<br>{lat}<br>{lon}",
                                icon=folium.Icon(color='red')).add_to(map_plot)

        # Adding legend for differentiating Inside and Outside coordinate points
        legend_html = """
            <div style="position: fixed; bottom:30px; right:30px; z-index:9999; font-size:11px;">
                <p><i class="fa fa-circle fa-1x" style="color:green"></i> Inside Boundary</p>
                <p><i class="fa fa-circle fa-1x" style="color:red"></i> Outside Boundary</p>
            </div>
            """
        map_plot.get_root().html.add_child(folium.Element(legend_html))

        # Adding various tiles option incoorporated within Folium library
        folium.TileLayer('CartoDB positron').add_to(map_plot) #Selected as Default
        folium.TileLayer('OpenStreetMap').add_to(map_plot)
        folium.TileLayer(tiles = 'https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}',
                attr = 'Esri',
                name = 'Esri Satellite',
                overlay = False,
                control = True
               ).add_to(map_plot)
        folium.TileLayer('esrinatgeoworldmap', name='Esri NatGeo WorldMap', attr=' Esri, Delorme, NAVTEQ').add_to(map_plot)

        # Full Screen option
        folium.plugins.Fullscreen().add_to(map_plot)

        # Layer control for each layers added can be checked & unchecked while selecting
        folium.LayerControl(position='topright',collapsed=True, autoZIndex=True).add_to(map_plot)

        # Saved the outputs
        map_plot.save(file_path)  
        
    def update_listbox(self):
        
        self.listbox3.clear()
        self.listbox4.clear()
        self.listbox5.clear()
        columns = self.data.columns.tolist()
        self.listbox3.addItems(columns)
        self.listbox4.addItems(columns)
        self.listbox5.addItems(columns)

    def standardize_cols(self):
        self.loc_col = self.listbox3.currentText()
        self.lat_col = self.listbox4.currentText()
        self.long_col = self.listbox5.currentText()
        self.append_text(f'Locnum Column selected: {self.loc_col}')
        self.append_text(f'Latitude Column selected: {self.lat_col}')
        self.append_text(f'Longitude Column selected: {self.long_col}')
        
    def geospatial_analysis(self):
        i = 0
        try:
            
            # get all the geometries and shapefiles
            gdf_low_res = gpd.read_file(self.shapefile_low_res)
            gdf_high_res = gpd.read_file(self.shapefile_high_res) 
            country_gdf_low_res = gdf_low_res.loc[gdf_low_res['NAME'] == self.country_name]
            country_geometry_low_res = country_gdf_low_res.geometry.iloc[0]
            country_geometry_buffer = country_geometry_low_res.buffer(-0.1)
            country_gdf_high_res = gdf_high_res.loc[gdf_high_res['NAME'] == self.country_name]
            country_geometry_high_res = country_gdf_high_res.geometry.iloc[0]
            minx, maxx, miny, maxy = gdf_low_res.iloc[0].geometry.bounds
            x_list = [minx, maxx]
            x_list.sort()
            y_list = [miny, maxy]
            y_list.sort()         
            
            self.append_text('Shapefiles successfully downloaded.')
            self.update_progress(20)
            
            # standardize the column names after being mapped
            self.data = self.data.rename(columns={f'{self.lat_col}': 'latitude', f'{self.long_col}': 'longitude', f'{self.loc_col}': 'ID'})      

            # add in a dataframe for the invalid data  
            nan_data = pd.DataFrame(columns=self.data.columns)

            # check if the row contains all necessary information
            column_names = self.data.columns.tolist()
                
            # initialise variables to store analysed outcomes
            inside = []
            inside_id = []
            outside_low_res = pd.DataFrame(columns=self.data.columns)
            outside_bounds = pd.DataFrame(columns=self.data.columns)
            outside_high_res = []
            inside_count = {}
            outside_count = {}
                
            i = 1

            for idx, row in self.data.iterrows():
                
                if row[column_names].isnull().any():
                    nan_data = pd.concat([nan_data, row.to_frame().T])
                    self.data.drop(idx, inplace=True)
                    print('row dropped due to null data')
                    continue
                
                # check if the correct data type has been inputted into the column
                try:
                    self.data.at[idx, 'latitude'] = float(row['latitude'])
                    self.data.at[idx, 'longitude'] = float(row['longitude'])
                except Exception:
                    nan_data = pd.concat([nan_data, row.to_frame().T])
                    self.data.drop(idx, inplace=True)
                    print('row dropped due to wrong data type')
                    continue
            
            for idx, row in self.data.iterrows():

                print(i)
                i += 1
                
                # quickly remove data in incompatible zones
                if x_list[0] <= row['longitude'] <= x_list[1] == False:
                    outside_bounds = pd.concat([outside_bounds, row.to_frame().T])
                    self.data.drop(idx, inplace=True)
                    print('row dropped due to not in bounds')
                if y_list[0] <= row['latitude'] <= y_list[1] == False:
                    outside_bounds = pd.concat([outside_bounds, row.to_frame().T])
                    self.data.drop(idx, inplace=True)
                    print('row dropped due to not in bounds')
                    continue
                
                location = Point(row['longitude'], row['latitude'])
                lat_long = (row['latitude'], row['longitude'])

                # add to out variables
                if location.within(country_geometry_buffer):
                    inside.append((row['ID'], row['latitude'], row['longitude']))
                    inside_count[lat_long] = inside_count.get(lat_long, 0) + 1
                    inside_id.append(row['ID'])
                else:
                    df_temporary = pd.DataFrame(data=[row], columns=outside_low_res.columns)
                    outside_low_res = pd.concat([outside_low_res, df_temporary], ignore_index=True)

            self.append_text('Geospatial analysis at low resolution finished.')

            for idx, row in outside_low_res.iterrows():
                location = Point(row['longitude'], row['latitude'])
                lat_long = (row['latitude'], row['longitude'])
                # add to our variables
                if location.within(country_geometry_high_res):
                    inside.append((row['ID'], row['latitude'], row['longitude']))
                    inside_count[lat_long] = inside_count.get(lat_long, 0) + 1
                    outside_low_res.drop(idx, inplace=True)
                else:
                    outside_high_res.append((row['ID'], row['latitude'], row['longitude']))
                    outside_count[lat_long] = outside_count.get(lat_long, 0) + 1
            
            outside_low_res = pd.concat([outside_low_res, outside_bounds])

            self.append_text('Geospatial analysis at high resolution finished.')
                                   
            self.update_progress(50)

            inside_df = self.data[self.data['ID'].isin(inside_id)]
            inside_file_path = self.file_path.replace('map_output.html', 'inside_LL.xlsx')
            invalid_LL_file_path = self.file_path.replace('map_output.html', 'invalid_LL.xlsx')
            outside_low_res.to_csv(invalid_LL_file_path, index=False)
            print('outside_excel_made')
            inside_df.to_csv(inside_file_path, index=False)
            print('inside_excel_made')
            self.append_text('Invalid LatLong excel file successfully created.')

            user_reporting_file_path = self.file_path.replace('map_output.html', 'user_reporting.txt')
            total = len(inside) + len(outside_high_res) + len(nan_data) 
            percentage = (len(outside_high_res) / total) * 100

            with open(user_reporting_file_path, "w") as file:
                file.write(f'Total points inside: {len(inside)} \n'
                           f'Total points outside: {len(outside_high_res)} \n'
                           f'Total incomplete data points: {len(nan_data)} \n'
                           f'Total points: {total} \n'
                           f'Percentage of data outside: {percentage:.1f}% \n')
            print('user_reporting_made')

            # Call the plot_maps method with the relevant data
            self.plot_maps(self.data, country_geometry_high_res, self.map_output, outside_high_res, inside_count, self.file_path, country_geometry_low_res, inside_df, outside_low_res)
            self.append_text('Map output successfully created.')
            self.append_text('Geospatial analysis completed.')

            self.update_progress(80) 
        
         
        except Exception as e:
            my_list = [self.data, self.shapefile_low_res, self.shapefile_high_res, self.country_name, self.map_output, self.file_path, self.lat_col, self.long_col, self.loc_col]
            counter = 0
            for i in my_list:
                if i is None:
                    counter += 1
                    break
                else:
                    continue
            if counter != 0:
                self.append_text(f'Not all fields were satisfied. Please ensure all inputs necessary have been inputted correctly')
            else:
                self.append_text(f'Error during geospatial analysis: {str(e)}')
                           
            
        self.update_progress(100)
        
    def reverse_geocode(self):
        if self.excel_path:
            invalid_LL_file_path = self.file_path.replace(f'map_output.html', 'invalid_LL.xlsx')
            if path.exists(invalid_LL_file_path):
                geolocator = Nominatim(user_agent="geoapiExercises")
                df = pd.read_excel(invalid_LL_file_path)
                locations = []
                for idx, row in df.iterrows():
                    try:
                        location = geolocator.reverse(f"{row['latitude']}, {row['longitude']}")
                        if location:
                            locations.append(location.address)
                        else:
                            locations.append("N/A")
                    except:
                        locations.append("N/A")
                df['Country'] = locations
                df.to_excel(invalid_LL_file_path, index=False)
                self.append_text(f'Reverse geocoding completed. Updated file: {invalid_LL_file_path}')
            else:
                self.append_text('Invalid_LL.xlsx file does not exist.')
        else:
            self.append_text('Please select an Excel file first.')
       
    def update_progress(self, progress_value):
        self.progress_bar.setValue(progress_value)
    
    def exit_app(self):
        self.close()
           
def main():
    
    app = QApplication(sys.argv)
    window = CoordinateValidationTool()
    return app.exec_()

if __name__ == '__main__':
    start_time = time.time()
    exit_code = main()
    end_time = time.time()
    print(end_time - start_time)
    sys.exit(exit_code)  
    
