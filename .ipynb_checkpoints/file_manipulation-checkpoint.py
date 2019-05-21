from tkinter import Tk
from tkinter.filedialog import askopenfilename, askdirectory, asksaveasfilename
import pandas as pd
import os

def get_file():
    win = Tk()
    win.withdraw()
    win.wm_attributes('-topmost', 1)
    filename = askopenfilename(parent=win)
    return filename

def get_folder():
    win = Tk()
    win.withdraw()
    win.wm_attributes('-topmost', 1)
    folder = askdirectory(parent=win)
    return folder

def saveas_filename():
    win = Tk()
    win.withdraw()
    win.wm_attributes('-topmost', 1)
    save_filename = asksaveasfilename(title='What would you like to name your excel file?',defaultextension = 'xlsx')
    return save_filename

def sort_data():
    #select a file
    filename = get_file() 
    #import data into a dataframe, the separator in this example is a tab (\t) and the headers are on row 9 (python is 0 indexed so row 1 in the text file is row 0 to python)
    data = pd.read_csv(filename, sep='\t', header=9) 
    #group the data you are interested in. In this example, each instance of the nuclei - ROI No column is summed to get the total number of nuclei
    extracted_data = data.groupby(['Row', 'Column', 'Field', 'Concentration', 'Cell Type', 'Replicate', 'Compound'], as_index=False)['Nuclei - ROI No'].sum()
    #rename the Nuclei - ROI No column to Nuclear Count
    extracted_data.rename(columns = {'Nuclei - ROI No':'Nuclear Count'}, inplace = True)
    #calculate the mean of the Nuclei grouped by Row and Column. 
#Grouping by Row and Column means that the mean of all the nuclei across fields within a single Row/Column (Well) will be calculated.
#rename the column that is created to Average Nuclear Count
    averaged_nuclei = extracted_data.groupby(['Row', 'Column', 'Concentration', 'Cell Type', 'Replicate', 'Compound'], as_index=False)['Nuclear Count'].mean().rename(columns={'Nuclear Count':'Average Nuclear Count'})
    #calculate the standard deviation of the Nuclei grouped by Row and Column. 
#Grouping by Row and Column means that the standard deviation of all the nuclei across fields within a single Row/Column (Well) will be calculated. 
#rename the column that is created to std
    std = extracted_data.groupby(['Row', 'Column'], as_index=False).std().rename(columns={'Nuclear Count':'std'})
    #add the standard deviation calculations to the main dataframe as a column called Standard Deviation
    averaged_nuclei['Standard Deviation'] = std['std'].tolist()
    #create a pivot table that shows a summary of the data (average nuclear count and standard deviation) as a function of compound and concentration 
    summary = pd.pivot_table(averaged_nuclei,index=["Compound","Concentration"], values=['Average Nuclear Count', 'Standard Deviation'])
    summary.reset_index(inplace=True)

    savefilename = saveas_filename()
    writer = pd.ExcelWriter(savefilename, engine='xlsxwriter')
    summary.to_excel(writer, sheet_name='summary')
    averaged_nuclei.to_excel(writer, sheet_name='average_nuclear_count')
    extracted_data.to_excel(writer, sheet_name='binned_data')
    data.to_excel(writer, sheet_name='raw_data')
    
    sheet4 = writer.sheets['raw_data']
    sheet3 = writer.sheets['binned_data']
    sheet2 = writer.sheets['average_nuclear_count']
    sheet1 = writer.sheets['summary']
    
    def get_column_width(dataframe):
        return [max([len(str(i))*1.25 for i in dataframe.index.values])]+[len(str(i))*1.25 for i in dataframe.columns]

    column_widths_4 = get_column_width(data)
    column_widths_3 = get_column_width(extracted_data)
    column_widths_2 = get_column_width(averaged_nuclei)
    column_widths_1 = get_column_width(summary)

    [sheet1.set_column(i,i,width) for i, width in enumerate(column_widths_1)]
    [sheet2.set_column(i,i,width) for i, width in enumerate(column_widths_2)]
    [sheet3.set_column(i,i,width) for i, width in enumerate(column_widths_3)]
    [sheet4.set_column(i,i,width) for i, width in enumerate(column_widths_4)];
    
    writer.save()
    
sort_data()