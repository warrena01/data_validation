import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QLabel, QFileDialog, QComboBox, QTextEdit,QProgressBar, QCheckBox
from PyQt5.QtCore import Qt
import geopandas as gpd
import pandas as pd
import folium
from shapely.geometry import Point
from folium.plugins import MarkerCluster
import glob
import openpyxl
from rtree.index import Index
import numpy
import time
import os

class CoordinateValidationTool(QMainWindow):
    
    def __init__(self):
        super().__init__()

        # set up attributes 
        self.data = None
        self.dir_path = None
        self.country_name = None
        self.gdf = None
        self.yesono = 'No'
        self.lat_col = None
        self.long_col = None
        self.cresta_col = None
        self.locnum_col = None

        # set up progress bar
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setGeometry(200, 490, 250, 30)
        self.progress_bar.setOrientation(Qt.Horizontal)
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0) 
        
        self.initUI()

    def initUI(self):

        # create the title and the geometry of the window
        self.setWindowTitle('Cresta Validation Tool')
        self.setGeometry(250, 250, 650, 650)

        '''GET THE DATA FROM AN EXCEL FILE'''
        self.label1 = QLabel('Step 1: Select the excel workbook', self)
        self.label1.setGeometry(30, 20, 400, 20)

        self.btn1 = QPushButton('Choose File', self)
        self.btn1.setGeometry(30, 50, 100, 30)
        self.btn1.clicked.connect(self.get_excel_doc)

        '''CHOOSE COUNTRY AND GET SHAPEFILE'''
        self.label2 = QLabel('Step 2: Select the country, then hit confirm', self)
        self.label2.setGeometry(30, 100, 400, 20)

        self.listbox1 = QComboBox(self)
        self.listbox1.setGeometry(30, 130, 200, 30)
        if os.path.exists(r'J:\cms\Internal\Territories\LAC\3. Vendor & Internal Models\LAC Geodata Validation Tools\LAC CRESTA ZONES PER COUNTRY - DO NOT MOVE\South American Cresta Zones Per Country High Res'):
            countries = glob.glob(r'J:\cms\Internal\Territories\LAC\3. Vendor & Internal Models\LAC Geodata Validation Tools\LAC CRESTA ZONES PER COUNTRY - DO NOT MOVE\South American Cresta Zones Per Country High Res\*\\')
        elif os.path.exists(r'J:\Analytics\cms\Internal\Territories\LAC\3. Vendor & Internal Models\LAC Geodata Validation Tools\LAC CRESTA ZONES PER COUNTRY - DO NOT MOVE\South American Cresta Zones Per Country High Res'):
            countries = glob.glob(r'J:\Analytics\cms\Internal\Territories\LAC\3. Vendor & Internal Models\LAC Geodata Validation Tools\LAC CRESTA ZONES PER COUNTRY - DO NOT MOVE\South American Cresta Zones Per Country High Res\*\\')
        countries_list = []
        for string in countries:
            split_parts = string.rsplit("\\", 2)
            country = split_parts[1]
            countries_list.append(country)
        self.listbox1.addItems(countries_list)
        self.listbox1.setCurrentIndex(-1)

        self.btn2 = QPushButton('Confirm', self)
        self.btn2.setGeometry(30, 160, 100, 30)
        self.btn2.clicked.connect(self.choose_country)

        '''CHECK FOR CRESTA DATA GIVEN'''
        self.label3 = QLabel('Step 3: Tick if crestas are given as a number', self)
        self.label3.setGeometry(30, 210, 500, 20)

        self.checkbox = QCheckBox("Enable Option", self)
        self.checkbox.setGeometry(30, 240, 20, 20) 
        self.checkbox.stateChanged.connect(self.cresta_data_given)

        '''MAP THE 4 NECESSARY COLUMNS'''
        self.label4 = QLabel('Step 4: Map your columns', self)
        self.label4.setGeometry(30, 290, 500, 20)

        self.label5 = QLabel('Locnum', self)
        self.label5.setGeometry(30, 320, 100, 30)
        self.listbox2 = QComboBox(self)
        self.listbox2.setGeometry(150, 320, 200, 30)
        self.listbox2.addItems([""])
        self.listbox2.setCurrentIndex(-1)

        self.label6 = QLabel('Latitude', self)
        self.label6.setGeometry(30, 350, 100, 30)
        self.listbox3 = QComboBox(self)
        self.listbox3.setGeometry(150, 350, 200, 30)
        self.listbox3.addItems([""])
        self.listbox3.setCurrentIndex(-1)

        self.label7 = QLabel('Longitude', self)
        self.label7.setGeometry(30, 380, 100, 30)
        self.listbox4 = QComboBox(self)
        self.listbox4.setGeometry(150, 380, 200, 30)
        self.listbox4.addItems([""])
        self.listbox4.setCurrentIndex(-1)

        self.label8 = QLabel('Cresta', self)
        self.label8.setGeometry(30, 410, 100, 30)
        self.listbox5 = QComboBox(self)
        self.listbox5.setGeometry(150, 410, 200, 30)
        self.listbox5.addItems([""])
        self.listbox5.setCurrentIndex(-1)

        self.btn3 = QPushButton('Confirm', self)
        self.btn3.setGeometry(150, 440, 100, 30)
        self.btn3.clicked.connect(self.standardize_cols)
        
        '''START ANALYSIS BUTTON'''
        self.btn4 = QPushButton('Start Analysis', self)
        self.btn4.setGeometry(30, 490, 130, 30)
        self.btn4.clicked.connect(self.geospatial_analysis)
        
        '''CREATE THE EXIT BUTTON'''
        self.exit_btn = QPushButton('Exit', self)
        self.exit_btn.setGeometry(30, 540, 100, 30)  
        self.exit_btn.clicked.connect(self.exit_app)
        
        '''CREATE THE USER REPORTING WINDOW'''
        self.textOutput = QTextEdit(self)
        self.textOutput.setGeometry(400, 30, 230, 380)
        self.textOutput.setReadOnly(True)

        self.show()
        
    def exit_app(self):
        self.close()

        
    def get_outputs(self, inside_id, cresta_ambiguity_id, outside_id, nan_data, inside_df, cresta_ambiguity):
        
        def CamelCase(word):
            words = word.split() 
            camel_case = ' '.join(word.capitalize() for word in words)
            return camel_case 
        
        cresta_ambiguity_df = self.data[self.data['LOCNUM_AJG'].isin(cresta_ambiguity_id)]
        outside_df = self.data[self.data['LOCNUM_AJG'].isin(outside_id)]
        
        if len(cresta_ambiguity_df) > 0:
            if self.yesono == 'Yes':
                for id, row in cresta_ambiguity_df.iterrows():
                    lat_long = (row['LAT_AJG'], row['LONG_AJG'])
                    value = cresta_ambiguity[(lat_long, row['CRESTA_AJG'])]['polygon_found']
                    cresta_ambiguity_df.at[id, 'CRESTA_FOUND_AJG'] = value.item() if isinstance(value, numpy.int64) else value
            else:
                for idx, row in cresta_ambiguity_df.iterrows():
                    lat_long = (row['LAT_AJG'], row['LONG_AJG'])
                    client_cresta_desc = CamelCase(row['CRESTA_AJG'])
                    client_cresta_num = self.gdf.loc[self.gdf['CRESTA_DES'] == client_cresta_desc, 'CRESTA_ID'].values[0]
                    cresta_ambiguity_df.loc[idx, 'CRESTA_NUM_FOUND_AJG'] = cresta_ambiguity[(lat_long, client_cresta_num)]['polygon_stated']
                    cresta_ambiguity_df.loc[idx, 'CRESTA_NUM_FOUND_AJG'] = cresta_ambiguity[(lat_long, client_cresta_num)]['polygon_found']
                    outside_df = self.data[self.data['LOCNUM_AJG'].isin(outside_id)]


        user_reporting_file_path = self.dir_path + (r'/User Reporting.txt')
        file = open(user_reporting_file_path, "w")
                
        file.write(f'Total correct cresta data points: {len(inside_id)} \n'
                   f'Total data points with inconsistent cresta/latlongs: {len(cresta_ambiguity_df)} \n'
                   f'Total data points outside any cresta zone: {len(outside_id)} \n'
                   f'Total valid points: {len(self.data)} \n'
                   f'Total points nan data: {len(nan_data)} \n')

        file.close()

        if len(inside_df) > 0:
            inside_df_path = self.dir_path + (r'/Correct Cresta Zone and Lat Longs.xlsx')
            inside_df.to_excel(inside_df_path, index=False)
        if len(cresta_ambiguity_df) > 0:
            cresta_ambiguity_df_path = self.dir_path + (r'/Inconsistent Cresta Zone and Lat Longs.xlsx')
            cresta_ambiguity_df.to_excel(cresta_ambiguity_df_path, index=False)
        if len(outside_df) > 0:
            outside_df_path = self.dir_path + (r'/Outside Any Cresta Zone Lat Longs.xlsx')
            outside_df.to_excel(outside_df_path, index=False)
        if len(nan_data) > 0:
            nan_data_path = self.dir_path + (r'/Incomplete Data.xlsx')
            nan_data.to_excel(nan_data_path, index=False)
        
        self.update_progress(100)
              
    def geospatial_analysis(self):

        def CamelCase(word):
            words = word.split() 
            camel_case = ' '.join(word.capitalize() for word in words)
            return camel_case  
        
        try:
        
            self.data = self.data.rename(columns={self.lat_col: 'LAT_AJG', self.long_col: 'LONG_AJG', self.cresta_col: 'CRESTA_AJG', self.locnum_col: 'LOCNUM_AJG'})
            nan_data = pd.DataFrame(columns=self.data.columns)
            
            # initialise variables
            inside = {}
            inside_id = []
            cresta_ambiguity = {}
            cresta_ambiguity_id = []
            outside = {}
            outside_id = []
            
            # create unified polygons
            gdf_filtered = self.gdf[['CRESTA_ID', 'geometry']]
            gdf_filtered =  gdf_filtered.dissolve(by='CRESTA_ID')
            gdf_filtered = gdf_filtered.reset_index()   

            # initialise a spatial index 
            my_spatial_index = Index()
            for id, country in gdf_filtered.iterrows():
                my_spatial_index.insert(id, country.geometry.bounds)

            '''CHECK FOR NULLS'''
            column_names = ['LAT_AJG', 'LONG_AJG', 'CRESTA_AJG', 'LOCNUM_AJG']
            for idx, row in self.data.iterrows():
                if row[column_names].isnull().any():
                    nan_data = pd.concat([nan_data, row.to_frame().T])
                    self.data.drop(idx, inplace=True)    
                    print('dropped due to null')
            
            '''STANDARDISE DR INPUTS'''
            if self.country_name == 'Dominican Republic':
                value_mapping = {'a': 1, 'b': 2, 'c': 3, 'd': 4, 'e': 5}
                for idx, row in self.data.iterrows():
                    if row['CRESTA_AJG'] in value_mapping:
                        self.data.at[idx, 'CRESTA_AJG'] = value_mapping[row['CRESTA_AJG']]

            if self.yesono == 'Yes':

                available_crestas = self.gdf['CRESTA_ID'].unique().tolist()
                available_crestas_int = list(map(int, available_crestas)) 

                # check each row of data
                for idx, row in self.data.iterrows():
                    
                    '''CHECK DATA IS A VALID TYPE'''
                    try:
                        self.data.at[idx, 'LAT_AJG'] = float(row['LAT_AJG'])
                        self.data.at[idx, 'LONG_AJG'] = float(row['LONG_AJG'])
                    except Exception:
                        nan_data = pd.concat([nan_data, row.to_frame().T])
                        self.data.drop(idx, inplace=True)
                        print('lat_long_wrong')
                        continue
                    try:
                        self.data.at[idx, 'CRESTA_AJG'] = int(row['CRESTA_AJG'])
                    except Exception:
                        nan_data = pd.concat([nan_data, row.to_frame().T])
                        self.data.drop(idx, inplace=True)
                        print('cresta_wrong')
                        continue
                        
                    '''CHECK DATA IS A VALID INPUT'''
                    
                    client_cresta_num = int(row['CRESTA_AJG'])
                    if client_cresta_num not in available_crestas_int:
                        nan_data = pd.concat([nan_data, row.to_frame().T]) 
                        self.data.drop(idx, inplace=True)
                        continue
                    
                    
                    point = Point(row['LONG_AJG'], row['LAT_AJG']) 
                    lat_long = (row['LAT_AJG'], row['LONG_AJG'])
                    # get the polygon for the given cresta
                    cresta_geometry = gdf_filtered.loc[gdf_filtered['CRESTA_ID'] == client_cresta_num, 'geometry'].values[0]

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
                                    cresta_ambiguity[key] = cresta_ambiguity.get(key, {'polygon_stated': client_cresta_num, 'polygon_found': gdf_filtered['CRESTA_ID'][row_id], 'count': 0})
                                    cresta_ambiguity[key]['count'] += 1
                                    cresta_ambiguity_id.append(row['LOCNUM_AJG'])
                                else: 
                                    continue      

                        else:
                            key = (lat_long, client_cresta_num)
                            outside[key] = outside.get(key, {'polygon_stated': client_cresta_num, 'count': 0})
                            outside[key]['count'] += 1
                            outside_id.append(row['LOCNUM_AJG'])
            
            if self.yesono == 'No':

                available_crestas = self.gdf['CRESTA_DES'].unique().tolist()

                # check each row of data
                for idx, row in self.data.iterrows():
                    
                    '''CHECK DATA IS A VALID TYPE'''
                    try:
                        self.data.at[idx, 'LAT_AJG'] = float(row['LAT_AJG'])
                        self.data.at[idx, 'LONG_AJG'] = float(row['LONG_AJG'])
                    except Exception:
                        nan_data = pd.concat([nan_data, row.to_frame().T])
                        self.data.drop(idx, inplace=True)
                        print('null data removed')
                        continue
                    
                    point = Point(row['LONG_AJG'], row['LAT_AJG']) 
                    lat_long = (row['LAT_AJG'], row['LONG_AJG']) 
                    # find the cresta name the client has given, standardize input
                    client_cresta_desc = CamelCase(row['CRESTA_AJG'])
                    
                    # check if the cresta aligns to a valid cresta zone
                    if client_cresta_desc not in available_crestas:
                        nan_data = pd.concat([nan_data, row.to_frame().T])
                        self.data.drop(idx, inplace=True)
                        continue
                    # find the cresta number this aligns to 
                    client_cresta_num = self.gdf.loc[self.gdf['CRESTA_DES'] == client_cresta_desc, 'CRESTA_ID'].values[0]
                    # get the polygon for given cresta
                    cresta_geometry = gdf_filtered.loc[gdf_filtered['CRESTA_ID'] == client_cresta_num, 'geometry'].values[0]

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
                                    cresta_ambiguity[key] = cresta_ambiguity.get(key, {'polygon_stated': client_cresta_num, 'polygon_found': gdf_filtered['CRESTA_ID'][row_id], 'count': 0})
                                    cresta_ambiguity[key]['count'] += 1
                                    cresta_ambiguity_id.append(row['LOCNUM_AJG'])
                                else: 
                                    continue      

                        else:
                            key = (lat_long, client_cresta_num)
                            outside[key] = outside.get(key, {'polygon_stated': client_cresta_num, 'count': 0})
                            outside[key]['count'] += 1
                            outside_id.append(row['LOCNUM_AJG'])

                    
            self.update_progress(50)
            self.append_text('Geospatial analysis complete.')            
            
            inside_df = self.data[self.data['LOCNUM_AJG'].isin(inside_id)]
            
            self.get_outputs(inside_id, cresta_ambiguity_id, outside_id, nan_data, inside_df, cresta_ambiguity)
        
            self.append_text('Outputs successfully created.')
            
        
        except Exception as e:
            # check everything has been filled in 
            my_list = [self.yesono, self.gdf, self.dir_path, self.data, self.lat_col, self.long_col, self.cresta_col, self.locnum_col]
            counter = 0
            for i in my_list:
                if (i is None):
                    counter += 1
                    break
                else:
                    continue
            if counter != 0:
                self.append_text(f'Not all fields were satisfied. Please ensure all inputs necessary have been inputted correctly')
            else:   
                self.append_text(f'Error during geospatial analysis: {str(e)}')
                 
    def get_excel_doc(self):
        self.excel_path, _ = QFileDialog.getOpenFileName(self, 'Select Excel File')
        if self.excel_path and self.excel_path.split(".")[-1] in ["xlsx", "xls"]:
            self.data = pd.read_excel(self.excel_path) 
            self.append_text('Excel file selected.')
            self.dir_path = self.excel_path.replace(self.excel_path.split("/")[-1], "")
            self.update_listbox()
            self.update_progress(5)
        elif self.excel_path and self.excel_path.split(".")[-1] in ["csv"]:
            self.data = pd.read_csv(self.excel_path, encoding='latin-1') 
            self.append_text('Excel file selected.')
            self.dir_path = self.excel_path.replace(self.excel_path.split("/")[-1], "")
            self.update_listbox()
            self.update_progress(5)
        else:
            self.append_text('Please choose a valid Excel file.')

    def choose_country(self):
        self.country_name = self.listbox1.currentText()
        self.append_text(f'Country selected: {self.country_name}')
        try:
            if os.path.exists(r'J:\cms\Internal\Territories\LAC\3. Vendor & Internal Models\LAC Geodata Validation Tools\LAC CRESTA ZONES PER COUNTRY - DO NOT MOVE\South American Cresta Zones Per Country High Res'):
                shp_path = rf'J:\cms\Internal\Territories\LAC\3. Vendor & Internal Models\LAC Geodata Validation Tools\LAC CRESTA ZONES PER COUNTRY - DO NOT MOVE\South American Cresta Zones Per Country High Res\{self.country_name}\{self.country_name}.shp'
                self.append_text(f'.shp download started')
                self.gdf = gpd.read_file(shp_path)
                self.append_text(f'Country successfully downloaded')
                self.update_progress(10)
            elif os.path.exists(r'J:\Analytics\cms\Internal\Territories\LAC\3. Vendor & Internal Models\LAC Geodata Validation Tools\LAC CRESTA ZONES PER COUNTRY - DO NOT MOVE\South American Cresta Zones Per Country High Res'):
                shp_path = rf'J:\Analytics\cms\Internal\Territories\LAC\3. Vendor & Internal Models\LAC Geodata Validation Tools\LAC CRESTA ZONES PER COUNTRY - DO NOT MOVE\South American Cresta Zones Per Country High Res\{self.country_name}\{self.country_name}.shp'
                self.append_text(f'.shp download started')
                self.gdf = gpd.read_file(shp_path)
                self.append_text(f'Country successfully downloaded')
                self.update_progress(10)
        except Exception:
            self.append_text('The file location appears to have moved. Ensure the following path file still exists "J:\cms\Internal\Territories\LAC\3. Vendor & Internal Models\LAC Geodata Validation Tools\LAC CRESTA ZONES PER COUNTRY - DO NOT MOVE\South American Cresta Zones Per Country High Res". Note, your file path may start with "Analytics"')
            
    def cresta_data_given(self, state):
        if state == 2:  # 2 represents the checked state
            self.append_text('Option selected: Cresta data has been given as a number')
            self.yesono = 'Yes'
        else:
            self.append_text('Option selected: Cresta data has not been given as a number')
            self.yesono = 'No'

    def update_listbox(self):
        
        self.listbox2.clear()
        self.listbox3.clear()
        self.listbox4.clear()
        self.listbox5.clear()

        columns = self.data.columns.tolist()
        self.listbox2.addItems(columns)
        self.listbox3.addItems(columns)
        self.listbox4.addItems(columns)
        self.listbox5.addItems(columns)

    def standardize_cols(self):
        self.lat_col = self.listbox3.currentText()
        self.long_col = self.listbox4.currentText()
        self.cresta_col = self.listbox5.currentText()
        self.locnum_col = self.listbox2.currentText()
        self.append_text(f'Locnum Column selected: {self.locnum_col}')
        self.append_text(f'Latitude Column selected: {self.lat_col}')
        self.append_text(f'Longitude Column selected: {self.long_col}')
        self.append_text(f'Cresta Column selected: {self.cresta_col}')

    def append_text(self, text):
        self.textOutput.append(text)
        self.textOutput.verticalScrollBar().setValue(self.textOutput.verticalScrollBar().maximum())
        self.textOutput.update()

    def update_progress(self, progress_value):
        self.progress_bar.setValue(progress_value)

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
