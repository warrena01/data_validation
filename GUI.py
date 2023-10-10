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
     
def get_shp():
    my_file_path = filedialog.askopenfilename()
    root.withdraw()
    print(my_file_path)

if __name__ == '__main__':
    window=tkinter.Tk()
    
    def get_text_input():
        text = textentry.get().upper()
        print(text)
    
    window.title('Coordinate Validation Tool')
    window.geometry('800x400')
    

    # add a title
    label = tkinter.Label(window, text='')
    label.pack()
    label = tkinter.Label(window, text='This is the Coordinate Validation Tool.')
    label.pack()
    label = tkinter.Label(window, text='')
    label.pack()
    
    # add a button get the user to select the excel file
    label = tkinter.Label(window, text='Select the excel workbook you want to use')
    label.pack()
    button = tkinter.Button(window, text='Find File', command = get_excel_doc)
    button.pack()
    label = tkinter.Label(window, text='')
    label.pack()
    
    # add a button get the user to select the shapefile
    label = tkinter.Label(window, text='Select the shapefile you want to use')
    label.pack()
    button = tkinter.Button(window, text='Find File', command = get_shp)
    button.pack()
    label = tkinter.Label(window, text='')
    label.pack()
    
    # add a button get the user input for country selected
    label = tkinter.Label(window, text='Enter the ISO A3 for the country being analysed, then hit confirm')
    label.pack()
    textentry = tkinter.Entry(window)
    textentry.pack()
    button = tkinter.Button(window, text='Confirm', command = get_text_input)
    button.pack()
    
    
    
    window.mainloop()
    
