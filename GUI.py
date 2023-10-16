def get_excel_doc():
    global data 
    root = tkinter.Tk()
    excel_doc = filedialog.askopenfilename()
    root.withdraw()
    if excel_doc.split(".")[-1] in [".xlsx", ".xls"]:
        data_1 = pd.read_excel(excel_doc)
        data_1.columns = ['ID', 'latitude', 'longitude']
        data = data_1
    else:
        messagebox.showerror(title="file type error", message="you chose the wrong file type, try again.")

def get_shp_1():
    global shapefile_low_res
    root = tkinter.Tk()
    my_file_path = filedialog.askopenfilename()
    root.withdraw()
    if my_file_path.split(".")[-1] in [".shp"]:
        shapefile_low_res = my_file_path
    else:
        messagebox.showerror(title="file type error", message="you chose the wrong file type, try again.")
        
def get_shp_2():
    global shapefile_high_res
    root = tkinter.Tk()
    my_file_path = filedialog.askopenfilename()
    root.withdraw()
    if my_file_path.split(".")[-1] in [".shp"]:
        shapefile_high_res = my_file_path
    else:
        messagebox.showerror(title="file type error", message="you chose the wrong file type, try again.")

def choose_country():
    global country_name, progress_value
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
    global file_name, task_completed
    text = textentry_file.get()
    file_name = text

def get_file_path():
    global file_path
    root = tkinter.Tk()
    my_file_path = filedialog.askdirectory()
    root.withdraw()
    my_file_path = str(my_file_path) + '/' + str(file_name) + '.html'
    file_path = my_file_path
    update_button_color()
    
def finished_tasks():
    global file_path
    if path.exists(file_path):
            messagebox.showinfo(title="task completed", message="All tasks are completed, you can now close all windows")
             
def geospatial_analysis():
    global data, shapefile_low_res, shapefile_high_res, country_name, map_output, file_name, file_path
    my_list = [data, shapefile_low_res, shapefile_high_res, country_name, map_output, file_name, file_path]
    for i in my_list:
        if i == None:
            messagebox.showerror(title="missing fields", message="you have not entered all fields, try again.")
            break
        break
    
def on_closing():
    window.quit()
    window.destroy()

if __name__ == '__main__':
    
    # create a window object
    window=tkinter.Tk()
    window.title('Coordinate Validation Tool')
    window.geometry('448x400')
    
   # Create a canvas that fills the entire window
    canvas = tkinter.Canvas(window)
    canvas.pack(side=tkinter.LEFT, fill=tkinter.BOTH, expand=True)
    frame = tkinter.Frame(canvas)
    canvas.create_window(3, 0, window=frame)
    def set_scroll_to_top(event):
        canvas.yview_moveto(0)
    canvas.bind("<Map>", set_scroll_to_top)

    # add a title
    label = tkinter.Label(frame, text='This is the Coordinate Validation Tool. Follow the Below Steps').pack(pady=5)

    # add a button to get the user to select the excel file
    label = tkinter.Label(frame, text='    Step 1: Select the excel workbook containing the ID, lat and long coordinates').pack()
    button = tkinter.Button(frame, text='Choose File', command = get_excel_doc).pack(pady=5)
    
    # add a button to get the user to select the shapefile at the low resolution
    label = tkinter.Label(frame, text='Step 2: Select the appropriate low resolution shapefile for your data').pack()
    button = tkinter.Button(frame, text='Choose File', command = get_shp_1).pack(pady=5)
        
    # add a button to get the user to select the shapefile at the high resolution
    label = tkinter.Label(frame, text='Step 3: Select the appropriate high resolution shapefile for your data').pack()
    button = tkinter.Button(frame, text='Choose File', command = get_shp_2).pack(pady=5)
    
    # add a listbox to get user to select country, add scrollbar and confirm button
    label = tkinter.Label(frame, text='Step 4: Select the country, then hit confirm').pack()
    listbox_frame = tkinter.Frame(frame)
    listbox_frame.pack()
    listvar = tkinter.StringVar(value=['Cuba', 'Bahamas', 'Bermuda', 'Brazil', 'Dominican Republic', 'Haiti', 'El Salvador', 'Guatemala', 'Costa Rica', 'Colombia', 'Turkey', 'Trinidad and Tobago', 'Grenada', 'Barbados', 'Saint Lucia', 'Dominica', 'Antigua and Barbuda', 'Saint Kitts and Nevis', 'Anguilla', 'Jamaica', 'Peru', 'Nicaragua', 'Argentina', 'Curacao', 'Aruba', 'U.S. Virgin Islands', 'Saint Barthelemy', 'Puerto Rico', 'Cayman Islands', 'Bolivia', 'Saint Martin', 'Suriname', 'Paraguay', 'Montserrat', 'Malta', 'Portugal'])
    listbox = tkinter.Listbox(listbox_frame, listvariable=listvar, height=5)
    listbox.grid(row=0, column=0)
    scrollbar = tkinter.Scrollbar(listbox_frame, orient=tkinter.VERTICAL, command=listbox.yview)
    scrollbar.grid(row=0, column=1, sticky=tkinter.NS)
    button = tkinter.Button(frame, text='Confirm', command=choose_country).pack(pady=5)

    # add a listbox to get user to select output, add scrollbar and confirm button
    label = tkinter.Label(frame, text='Step 5: Select the output map, then hit confirm').pack()
    listbox_frame2 = tkinter.Frame(frame)
    listbox_frame2.pack()
    listvar2 = tkinter.StringVar(listbox_frame2, value= ['Heatmap', 'Heatmap Overlap', 'Point Data Map', 'All'])
    listbox_frame = tkinter.Frame(frame)
    listbox_frame.pack()
    listbox2 = tkinter.Listbox(listbox_frame, listvariable=listvar2, height=4)
    listbox2.pack()
    button = tkinter.Button(frame, text='Confirm', command=choose_output)
    button.pack(pady=5)
    
    # add a button to get the desired file name
    label = tkinter.Label(frame, text='Step 6: Choose the name you want to save the file in').pack()
    textentry_file = tkinter.Entry(frame).pack()
    button = tkinter.Button(frame, text='Confirm', command = get_file_name).pack(pady=5)
    
    # add a button get the user to select the destination folder
    label = tkinter.Label(frame, text='Step 7: Select the folder you would like your outputs saved to').pack()
    button = tkinter.Button(frame, text='Choose Folder', command = get_file_path).pack(pady=5)
    
    # start analysis
    label = tkinter.Label(frame, text='Step 8: Start Analysis').pack()
    button = tkinter.Button(frame, text='Go', command = geospatial_analysis).pack()
    label = tkinter.Label(frame, text='').pack(pady=5)

    # Add a scrollbar to the window
    def configure_scroll_region(event):
        canvas.configure(scrollregion=canvas.bbox("all"))
    frame.bind("<Configure>", configure_scroll_region)
    scrollbar = tkinter.Scrollbar(window, orient=tkinter.VERTICAL, command=canvas.yview)
    scrollbar.pack(side=tkinter.RIGHT, fill=tkinter.Y)
    canvas.configure(yscrollcommand=scrollbar.set)
    
    window.protocol("WM_DELETE_WINDOW", on_closing)
    window.mainloop()
