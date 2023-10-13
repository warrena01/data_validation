import tkinter
from tkinter import filedialog
import geopandas as gpd
import pandas as pd
import folium
from shapely.geometry import Point
import openpyxl
from folium.plugins import HeatMap, MarkerCluster

# intiate global variables
data = None # this will hold the excel workfile
shapefile_low_res = None # this will hold the shapefile at a low resolution
shapefile_high_res = None # this will hold the shapefile at a high resolution
country_name = None # this will hold the variable that names the country 
map_output = None # this will hold the map output wanted 
file_name = None # this will hold the file name they want to save it as 
file_path = None # this will hold the path link to the new file

def get_excel_doc():
    global data 
    root = tkinter.Tk()
    excel_doc = filedialog.askopenfilename()
    file_types = [("Excel Files", "*.xlsx;*.xls"), ("All Files", "*.*")]
    root.withdraw()
    data_1 = pd.read_excel(excel_doc)
    data_1.columns = ['ID', 'latitude', 'longitude']
    data = data_1
    
def get_shp_1():
    global shapefile_low_res
    root = tkinter.Tk()
    my_file_path = filedialog.askopenfilename()
    file_types = [("Shapefile Files", "*.shp"), ("All Files", "*.*")]
    root.withdraw()
    shapefile_low_res = my_file_path
    
def get_shp_2():
    global shapefile_high_res
    root = tkinter.Tk()
    my_file_path = filedialog.askopenfilename()
    file_types = [("Shapefile Files", "*.shp"), ("All Files", "*.*")]
    root.withdraw()
    shapefile_high_res = my_file_path

def choose_country():
    global country_name
    selected_indices = listbox.curselection()
    if selected_indices:
        index = selected_indices[0]  # Assuming you expect only one selection
        country_name = listbox.get(index)
        
def choose_output():
    global map_output
    selected_indices = listbox2.curselection()
    if selected_indices:
        index = selected_indices[0]  # Assuming you expect only one selection
        map_output = listbox2.get(index)
        
    
def get_file_name():
    global file_name
    text = textentry_file.get()
    file_name = text
    
def get_file_path():
    global file_path
    root = tkinter.Tk()
    my_file_path = filedialog.askdirectory()
    root.withdraw()
    my_file_path = str(my_file_path) + '/' + str(file_name) + '.html'
    file_path = my_file_path

def geospatial_analysis():
    print(hello)
    
def on_closing():
    window.quit()
    window.destroy()

if __name__ == '__main__':
    window=tkinter.Tk()
    window.title('Coordinate Validation Tool')
    window.geometry('700x400')
    
    scrollbar2 = tkinter.Scrollbar(window)
    scrollbar2.pack(side=tkinter.RIGHT, fill=tkinter.Y)
    
    
    # add a title
    label = tkinter.Label(window, text='')
    label.pack()
    label = tkinter.Label(window, text='This is the Coordinate Validation Tool. Follow the Below Steps')
    label.pack()
    label = tkinter.Label(window, text='')
    label.pack()
    
    # add a button to get the user to select the excel file
    label = tkinter.Label(window, text='Step 1: Select the excel workbook containing the ID, lat and long coordinates')
    label.pack()
    button = tkinter.Button(window, text='Choose File', command = get_excel_doc)
    button.pack()
    label = tkinter.Label(window, text='')
    label.pack()
    
    # add a button to get the user to select the shapefile at the low resolution
    label = tkinter.Label(window, text='Step 2: Select the appropriate low resolution shapefile for your data')
    label.pack()
    button = tkinter.Button(window, text='Choose File', command = get_shp_1)
    button.pack()
    label = tkinter.Label(window, text='')
    label.pack()
    
    # add a button to get the user to select the shapefile at the high resolution
    label = tkinter.Label(window, text='Step 3: Select the appropriate high resolution shapefile for your data')
    label.pack()
    button = tkinter.Button(window, text='Choose File', command = get_shp_2)
    button.pack()
    label = tkinter.Label(window, text='')
    label.pack()

    # add a button to get the user input for country selected
    label = tkinter.Label(window, text='Step 4: Type the name of the country, then hit confirm')
    label.pack()
    listbox_frame = tkinter.Frame(window)
    listbox_frame.pack()
    listvar = tkinter.StringVar(listbox_frame, value= ['Cuba', 'Bahamas', 'Bermuda', 'Brazil', 'Dominican Republic', 'Haiti', 'El Salvador', 'Guatemala', 'Costa Rica', 'Colombia', 'Turkey', 'Trinidad and Tobago', 'Grenada', 'Barbados', 'Saint Lucia', 'Dominica', 'Antigua and Barbuda', 'Saint Kitts and Nevis', 'Anguilla', 'Jamaica', 'Peru', 'Nicaragua', 'Argentina', 'Curacao', 'Aruba', 'U.S. Virgin Islands', 'Saint Barthelemy', 'Puerto Rico', 'Cayman Islands', 'Bolivia', 'Saint Martin', 'Suriname', 'Paraguay', 'Montserrat', 'Malta', 'Portugal']) 
    listbox_frame = tkinter.Frame(window)
    listbox_frame.pack()
    listbox = tkinter.Listbox(listbox_frame, listvariable=listvar, height=5)
    listbox.grid(row=0, column=0)
    scrollbar = tkinter.Scrollbar(listbox_frame, orient=tkinter.VERTICAL)
    scrollbar.grid(row=0, column=1, sticky=tkinter.NS)
    listbox.config(yscrollcommand=scrollbar.set)
    scrollbar.config(command=listbox.yview)
    button = tkinter.Button(window, text='Confirm', command=choose_country)
    button.pack()
    label = tkinter.Label(window, text='')
    label.pack()
     
    # add a button to get the user input for map output desired
    label = tkinter.Label(window, text='Step 5: Choose output map')
    label.pack()
    listbox_frame2 = tkinter.Frame(window)
    listbox_frame2.pack()
    listvar2 = tkinter.StringVar(listbox_frame2, value= ['Heatmap', 'Heatmap Overlap', 'Point Data Map', 'All'])
    listbox_frame = tkinter.Frame(window)
    listbox_frame.pack()
    listbox2 = tkinter.Listbox(listbox_frame, listvariable=listvar2, height=4)
    listbox2.pack()
    button = tkinter.Button(window, text='Confirm', command=choose_output)
    button.pack()
    
    # add a button to get the desired file name
    label = tkinter.Label(window, text='Step 6: Choose the name you want to save the file in')
    label.pack()
    textentry_file = tkinter.Entry(window)
    textentry_file.pack()
    button = tkinter.Button(window, text='Confirm', command = get_file_name)
    button.pack()
    label = tkinter.Label(window, text='')
    label.pack()
    
    # add a button get the user to select the destination folder
    label = tkinter.Label(window, text='Step 7: Select the folder you would like your map saved to')
    label.pack()
    button = tkinter.Button(window, text='Choose Folder', command = get_file_path)
    button.pack()
    label = tkinter.Label(window, text='')
    label.pack()
    
    # add a go button to get the desired file name
    label = tkinter.Label(window, text='Step 8: Start Analysis')
    label.pack()
    button = tkinter.Button(window, text='Go', command = geospatial_analysis)
    button.pack()
    label = tkinter.Label(window, text='')
    label.pack()
    
    
    
    window.protocol("WM_DELETE_WINDOW", on_closing)
    window.mainloop()
