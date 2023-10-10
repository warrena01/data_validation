import tkinter
from tkinter import filedialog, PhotoImage
import pandas as pd
from PIL import Image, ImageTk

def get_excel_doc():
    root = tkinter.Tk()
    excel_doc = filedialog.askopenfilename()
    root.withdraw()
    data = pd.read_excel(excel_doc)
    data.columns = ['ID', 'latitude', 'longitude']
    
def high_res_only():
    print(checkvar.get())
     
def get_shp():
    my_file_path = filedialog.askopenfilename()
    root.withdraw()
    print(my_file_path)
    
def get_text_input():
    text = textentry.get().upper()
    print(text)

if __name__ == '__main__':
    window=tkinter.Tk()
    
   
    
    window.title('Coordinate Validation Tool')
    window.geometry('800x400')
    
    
   
    # add a button get the user to select the excel file
    label = tkinter.Label(window, text='Select the excel workbook you want to use')
    label.grid(row=1, column=1)
    button = tkinter.Button(window, text='Find File', command = get_excel_doc)
    button.grid(row=1, column=2)

    
    # add a button get the user to select the shapefile
    label = tkinter.Label(window, text='Select the shapefile you want to use')
    label.grid(row=3, column=1)
    button = tkinter.Button(window, text='Find File', command = get_shp)
    button.grid(row=3, column=2)

    
    # add a button get the user input for country selected
    label = tkinter.Label(window, text='Enter the ISO A3 for the country being analysed, then hit confirm')
    label.grid(row=5, column=1)
    textentry = tkinter.Entry(window)
    textentry.grid(row=5, column=2)
    button = tkinter.Button(window, text='Confirm', command = get_text_input)
    button.grid(row=5, column=3)
    
    # add a check button to check which analysis:
    checkvar = tkinter.StringVar()
    check = tkinter.Checkbutton(window, text='check me', variable=checkvar, onvalue='True', offvalue='False', command = high_res_only)
    check.grid(row=7, column = 2)
    
    
    
    window.mainloop()
    
    
    window.mainloop()
    
