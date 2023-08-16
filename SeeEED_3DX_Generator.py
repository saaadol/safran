import pandas as pd
import tkinter as tk
from tkinter.filedialog import askdirectory
from tkinter import filedialog, messagebox, ttk, HORIZONTAL
import warnings
import math
from PIL import ImageTk, Image
from ttkthemes import ThemedStyle
import  time
import threading
import time
import os

current_directory = os.getcwd()

def close_window():
    window.destroy()


def on_close():
    # Show the message box when the user clicks the close button (X)
    result = messagebox.askquestion("Exit", "Are you sure you want to exit?")
    if result == "yes":
        window.destroy()
        exit()


def select_airbus():
    window.aircraft_type = 'Airbus'
    # window.destroy()


# Function to handle the selection of Dassault option

def select_dassault():
    window.aircraft_type = 'Dassault'
    # window.destroy()


# Ignore all warnings
warnings.filterwarnings("ignore")


def remove_duplicates(df):
    index = []
    for i in range(1, len(df) - 1):
        if df.iloc[i, 1] == df.iloc[i + 1, 1]:
            index.append(i + 1)
    for y in range(len(index)):
        df.drop(index[y], inplace=True)


def remove_dup(df):
    pair = []
    output_df = pd.DataFrame(columns=df.columns)

    for index, row in df.iterrows():
        product_family = row['Product_Family']
        part_no = row['PartNo']

        # Check if the pair exists in the unique_pairs list
        if not pair:
            pair = (product_family, part_no)
            output_df = output_df.append(row)

        elif product_family == pair[0] and part_no == pair[1]:
            # pair = (product_family, part_no)
            # unique_pairs.append(pair)
            output_df = output_df.append(row)

        # elif (product_family, part_no) not in unique_pairs and (all(part_no != pair[1] and product_family != pair[0] for pair in unique_pairs) or any(part_no != pair[1] and product_family == pair[0] for pair in unique_pairs)):
        elif (product_family != pair[0] and part_no != pair[1]) or (product_family == pair[0] and part_no != pair[1]):
            # Create a tuple representing the pair of 'Product Family' and 'Part No'
            pair = (product_family, part_no)

            # Add the pair to the unique_pairs list
            # unique_pairs.append(pair)

            # Add the current row to the output DataFrame
            output_df = output_df.append(row)

        # elif all(part_no == pair[1] and product_family != pair[0] for pair in unique_pairs):
        # continue

    return output_df


def filter_cells(df, row_index, row, i):
    if window.aircraft_type == 'Airbus':
        product_family = ['A350-IEV', 'A320FR', 'A320GE', 'A320UK', 'A330FR', 'A330GE', 'A330UK', 'A380FR', 'A350',
                          'A380GE', 'A380UK', 'A400M', 'AIRBUS', 'ATR']
    elif window.aircraft_type == 'Dassault':
        product_family = ['ATL2', 'BUS-DASSAULT', 'DASSAULT', 'DKIT', 'F10X', 'F2000X', 'F50SURMAR', 'F5X', 'F6X',
                          'F7X', 'F8X', 'F900X', 'FALCON', 'M2000', 'M5000', 'MIRAGE', 'NEURON', 'RAFALE']
    else:
        window.aircraft_type = None
        print("Invalid aircraft type selected.")
        exit()
    df = df[df.iloc[:, 0].isin(product_family)]
    result = {}
    try:
        value = row.iloc[1]
        if len(str(row.iloc[7])) > 2:
            col_7_value = str(row.iloc[7])[:2]
        else:
            col_7_value = row.iloc[7]
        col_8_value = row.iloc[8]
    except IndexError:
        print(f"Row index '{row_index}' is out of range.")
        return {}

    if pd.notnull(col_7_value) and pd.notnull(col_8_value):
        if not (df[(df.iloc[:, 3] == col_7_value) & (df.iloc[:, 4] == col_8_value)]).empty:
            filtered_df = df[(df.iloc[:, 3] == col_7_value) & (df.iloc[:, 4] == col_8_value)]
            filtered_df = filtered_df.reset_index(drop=True)
            line = filtered_df.iloc[0, :]
            result[value] = line.iloc[i]
        else:
            filtered_df2 = df[df.iloc[:, 1] == value]
            filtered_df2 = filtered_df2.reset_index(drop=True)
            Cable_type = filtered_df2.iloc[0, 3]
            Cable_Gauge = filtered_df2.iloc[0, 4].astype(int)
            line = filtered_df2.iloc[0, :]
            if i == 1:
                result[value] = str(Cable_type) + str(Cable_Gauge) + 'Wire'
            else:
                result[value] = line.iloc[i]

    elif pd.isnull(col_7_value) or pd.isnull(col_8_value):
        filtered_df2 = df[df.iloc[:, 1] == value]
        filtered_df2 = filtered_df2.reset_index(drop=True)
        Cable_type = filtered_df2.iloc[0, 3]
        Cable_Gauge = filtered_df2.iloc[0, 4]
        line = filtered_df2.iloc[0, :]
        if i == 1:
            if pd.isnull(col_8_value):
                result[value] = str(Cable_type) + 'Core'
            else:
                result[value] = str(Cable_type) + str(int(Cable_Gauge)) if not math.isnan(Cable_Gauge) else ''
        else:
            result[value] = line.iloc[i]
    return result[value]


def remove_slash_C(df_output):
    df_output['Parent Name'] = df_output['Parent Name'].str.replace("/C", '')
    df_output['Name'] = df_output['Name'].str.replace("/C", '')
    df_output['(R) Title'] = df_output['(R) Title'].str.replace("/C", '')
    return df_output


color_mapping = {
    'Blue': 'B',
    'Red': 'R',
    'Yellow': 'Y',
    'Blue Light': 'BL',
    'Green': 'G',
    'Red and Blue': 'RB',
    'Yellow and Green': 'YG',
    'White': 'W',
    'Brown': 'Br',
    'Beige': 'Be',
    'Orange': 'O',
    'White and Blue': 'WB',
    'White and Orange': 'WO',
    'White and Brown': 'WBr',
    'Pink': 'Pk',
    'Violet': 'V',
    'White and Green': 'WG',
    'White and Black': 'WBL',
    'White and Red': 'WR',
    'Ligt Green': 'LG',
    'Light BLue': 'LB',
    'Dark Green': 'DG'
}


def concatenate_color(df1, row, index):
    color = row.iloc[45]
    return str(filter_cells(df1, index, row, 1)) + color_mapping.get(color, '')

    
class FileManager():
    def __init__(self, input_file, output_file, window):
        self.input_file =  pd.read_excel(input_file)
        self.output_file = output_file
        self.window = window

    def creating_progress_bar(self):
        self.progress_bar = ttk.Progressbar(self.window, orient=tk.HORIZONTAL, length=150, mode='indeterminate')
        self.progress_bar.start(8)
        self.progress_bar.place(relx=0.485, rely=0.35, anchor=tk.CENTER)
    def executing_threads(self):
        thread = threading.Thread(target=self.creating_progress_bar)
        thread.start()      

    def process_file(self):
        # Create the output dataframe with the desired headers
        output_columns = ['Parent Name', 'Name', 'Type', '(R) Title', '(R) description',
                          '(R) Sub-Type', '(R) Bend Radius', '(R) Linear Mass',
                          '(R) Outside Diameter', '(R) Color', 'Program']
        

        df_output = pd.DataFrame(columns=output_columns)

        # Filter rows based on the selected aircraft type

        if window.aircraft_type == 'Airbus':
            product_family = ['A350-IEV', 'A320FR', 'A320GE', 'A320UK', 'A330FR', 'A330GE', 'A330UK', 'A380FR', 'A350',
                              'A380GE', 'A380UK', 'A400M', 'AIRBUS', 'ATR']
        elif window.aircraft_type == 'Dassault':
            product_family = ['ATL2', 'BUS-DASSAULT', 'DASSAULT', 'DKIT', 'F10X', 'F2000X', 'F50SURMAR', 'F5X', 'F6X',
                              'F7X', 'F8X', 'F900X', 'FALCON', 'M2000', 'M5000', 'MIRAGE', 'NEURON', 'RAFALE']
        else:
            window.aircraft_type = None
            print("Invalid aircraft type selected.")
            exit()

        
        # Filter rows where column 6 values = 1 and column 7 values equal 'UNSHIELDED'
        filtered_data = self.input_file[
            (self.input_file.iloc[:, 0].isin(product_family)) & (self.input_file.iloc[:, 9] == 'Released') & (
                    self.input_file.iloc[:, 5] == 1) & (
                    self.input_file.iloc[:, 6] == 'UNSHIELDED')]

        filtered_data = filtered_data.sort_values(by=['PartNo', 'Product_Family'], ascending=True)
        #filtered_data = filtered_data.reset_index(drop=True)
        #filtered_data = filtered_data.drop_duplicates(subset=['PartNo'])
        filtered_data = filtered_data.reset_index(drop=True)

        # Map the data from the input to the output dataframe
        for _, row in filtered_data.iterrows():
            first_row = {
                'Parent Name': '',
                'Name': 'Root',
                'Type': 'Electrical Conductor Group',
                '(R) Title': '',
                '(R) description': '',
                '(R) Sub-Type': '',
                '(R) Bend Radius': '',
                '(R) Linear Mass': '',
                '(R) Outside Diameter': '',
                '(R) Color': '',
                'Program': ''
            }
            if _ == 0:
                df_output = df_output.append(first_row, ignore_index=True)

            else:
                additional_row1 = {
                    'Parent Name': 'Root',
                    'Name': row['PartNo'],
                    'Type': 'Electrical Conductor',
                    '(R) Title': row['PartNo'],
                    '(R) description': row['FR_Description'],
                    '(R) Sub-Type': '',
                    '(R) Bend Radius': row['Bend_Radius'],
                    '(R) Linear Mass': row['Linear_Weight__Kg_'],
                    '(R) Outside Diameter': row['External_Max_Diameter__M_'],
                    '(R) Color': row['Cable_Color'],
                    'Program': row['Product_Family']
                }
                additional_row2 = {
                    'Parent Name': 'Root',
                    'Name': row['PartNo'],
                    'Type': 'Electrical Conductor',
                    '(R) Title': str(row['Wire_Type']) + str(int(row['Wire_Gauge'])) if not math.isnan(
                        row['Wire_Gauge']) else str(row['Wire_Type']),
                    '(R) description': row['FR_Description'],
                    '(R) Sub-Type': '',
                    '(R) Bend Radius': row['Bend_Radius'],
                    '(R) Linear Mass': row['Linear_Weight__Kg_'],
                    '(R) Outside Diameter': row['External_Max_Diameter__M_'],
                    '(R) Color': row['Cable_Color'],
                    'Program': row['Product_Family']
                }
                if window.aircraft_type == 'Airbus':
                    df_output = df_output.append(additional_row2, ignore_index=True)
                elif window.aircraft_type == 'Dassault':
                    df_output = df_output.append(additional_row1, ignore_index=True)
        
        df_output = remove_slash_C(df_output)
        df_output = df_output.drop_duplicates(subset='Name')
        df_output = df_output.reset_index(drop=True)

        filtered_data2 = self.input_file[
            (self.input_file.iloc[:, 0].isin(product_family)) & (self.input_file.iloc[:, 9] == 'Released') &
            (self.input_file.iloc[:, 6] != 'UNSHIELDED')]

        filtered_data2 = filtered_data2.sort_values(by=['PartNo', 'Product_Family'], ascending=True)
        filtered_data2 = filtered_data2.reset_index(drop=True)
        filtered_data2 = remove_dup(filtered_data2)
        filtered_data2 = filtered_data2.sort_values(by=['PartNo', 'Product_Family'], ascending=True)
        filtered_data2 = filtered_data2.reset_index(drop=True)

        for _, row in filtered_data2.iterrows():
            wire_type = row['Shield']

            additional_row1 = {
                'Parent Name': 'Root',
                'Name': row['PartNo'],
                'Type': 'Electrical Conductor Group',
                '(R) Title': row['PartNo'],
                '(R) description': row['FR_Description'],
                '(R) Sub-Type': '',
                '(R) Bend Radius': row['Bend_Radius'],
                '(R) Linear Mass': row['Linear_Weight__Kg_'],
                '(R) Outside Diameter': row['External_Max_Diameter__M_'],
                '(R) Color': row['Cable_Color'],
                'Program': row['Product_Family']
            }
            additional_row2 = {
                'Parent Name': row['PartNo'],
                'Name': filter_cells(self.input_file, _, row, 1) + str(row['Wire_Color']),
                # concatenate_color(self.input_file,row,_),
                'Type': 'Electrical Conductor',
                '(R) Title': filter_cells(self.input_file, _, row, 1),
                '(R) description': filter_cells(self.input_file, _, row, 12),
                '(R) Sub-Type': row['Cable_Name'],
                '(R) Bend Radius': filter_cells(self.input_file, _, row, 17),
                '(R) Linear Mass': filter_cells(self.input_file, _, row, 24),
                '(R) Outside Diameter': filter_cells(self.input_file, _, row, 21),
                '(R) Color': '',
                'Program': row['Product_Family']
            }

            additional_row3 = {
                'Parent Name': row['PartNo'],
                'Name': str(row['PartNo']) + 'SH',
                'Type': 'Electrical Conductor',
                '(R) Title': str(row['PartNo']) + 'SH',
                '(R) description': 'BLINDAGE - SH',
                '(R) Sub-Type': '',
                '(R) Bend Radius': '',
                '(R) Linear Mass': '',
                '(R) Outside Diameter': '',
                '(R) Color': '',
                'Program': row['Product_Family']
            }
            additional_row4 = {
                'Parent Name': row['PartNo'],
                'Name': str(row['PartNo']) + 'SH1',
                'Type': 'Electrical Conductor',
                '(R) Title': str(row['PartNo']) + 'SH1',
                '(R) description': 'BLINDAGE - SH',
                '(R) Sub-Type': '',
                '(R) Bend Radius': '',
                '(R) Linear Mass': '',
                '(R) Outside Diameter': '',
                '(R) Color': '',
                'Program': row['Product_Family']
            }
            additional_row5 = {
                'Parent Name': row['PartNo'],
                'Name': str(row['PartNo']) + 'SH2',
                'Type': 'Electrical Conductor',
                '(R) Title': str(row['PartNo']) + 'SH2',
                '(R) description': 'BLINDAGE - SH',
                '(R) Sub-Type': '',
                '(R) Bend Radius': '',
                '(R) Linear Mass': '',
                '(R) Outside Diameter': '',
                '(R) Color': '',
                'Program': row['Product_Family']
            }
            additional_row6 = {
                'Parent Name': row['PartNo'],
                'Name': str(row['PartNo']) + 'SH3',
                'Type': 'Electrical Conductor',
                '(R) Title': str(row['PartNo']) + 'SH3',
                '(R) description': 'BLINDAGE - SH',
                '(R) Sub-Type': '',
                '(R) Bend Radius': '',
                '(R) Linear Mass': '',
                '(R) Outside Diameter': '',
                '(R) Color': '',
                'Program': row['Product_Family']
            }
            
            var = _ - 1
            next = _ + 1
            flag = 0
            wire_gauge = str(int(row['Wire_Gauge'])) if not math.isnan(row['Wire_Gauge']) else ''
            cable_gauge = str(int(row['Cable_Gauge'])) if not math.isnan(row['Cable_Gauge']) else ''
            if _ == 0:
                if window.aircraft_type == 'Airbus':
                    additional_row2['(R) Title'] = str(row['Wire_Type']) + wire_gauge
                    additional_row1['(R) Title'] = str(row['Cable_Type']) + cable_gauge
                    df_output = df_output.append(additional_row1, ignore_index=True)
                    df_output = df_output.append(additional_row2, ignore_index=True)
                else:
                    df_output = df_output.append(additional_row1, ignore_index=True)
                    df_output = df_output.append(additional_row2, ignore_index=True)
            else:
                current_product_family = row['Product_Family']
                if next >= len(filtered_data2):
                    flag = 1

                elif var < len(filtered_data2) and next < len(filtered_data2):

                    prev_product_family = filtered_data2.loc[var, 'Product_Family']
                    next_product_family = filtered_data2.loc[next, 'Product_Family']
                    prev_part_no = filtered_data2.loc[var, 'PartNo']
                    next_part_no = filtered_data2.loc[next, 'PartNo']
                else:
                    prev_product_family = None
                    prev_part_no = None
                    next_part_no = None
                    next_product_family = None

                if current_product_family != prev_product_family and next < len(filtered_data2):
                    if window.aircraft_type == 'Airbus':
                        additional_row1['(R) Title'] = str(row['Cable_Type']) + cable_gauge
                        additional_row2['(R) Title'] = str(row['Wire_Type']) + wire_gauge
                        df_output = df_output.append(additional_row1, ignore_index=True)
                        df_output = df_output.append(additional_row2, ignore_index=True)
                    else:
                        df_output = df_output.append(additional_row1, ignore_index=True)
                        df_output = df_output.append(additional_row2, ignore_index=True)

                else:
                    if row['PartNo'] == prev_part_no or flag == 1:
                        if window.aircraft_type == 'Airbus':
                            additional_row2['(R) Title'] = str(row['Wire_Type']) + wire_gauge
                            df_output = df_output.append(additional_row2, ignore_index=True)
                        else:
                            df_output = df_output.append(additional_row2, ignore_index=True)

                    else:
                        if window.aircraft_type == 'Airbus':
                            additional_row1['(R) Title'] = str(row['Cable_Type']) + cable_gauge
                            additional_row2['(R) Title'] = str(row['Wire_Type']) + wire_gauge
                            df_output = df_output.append(additional_row1, ignore_index=True)
                            df_output = df_output.append(additional_row2, ignore_index=True)
                        else:
                            df_output = df_output.append(additional_row1, ignore_index=True)
                            df_output = df_output.append(additional_row2, ignore_index=True)
                if current_product_family != next_product_family or row['PartNo'] != next_part_no or flag == 1:
                    if wire_type == 'SHIELDED':
                        if window.aircraft_type == 'Airbus':
                            additional_row3['(R) Title'] = str(row['Cable_Type']) + cable_gauge + 'SH'
                            df_output = df_output.append(additional_row3, ignore_index=True)
                        else:
                            df_output = df_output.append(additional_row3, ignore_index=True)
                    elif wire_type == 'DOUBLE SHIELD':
                        for k in range(2):
                            if k == 0:
                                if window.aircraft_type == 'Airbus':
                                    additional_row4['(R) Title'] = str(row['Cable_Type']) + cable_gauge + 'SH1'
                                    df_output = df_output.append(additional_row4, ignore_index=True)
                                else:
                                    df_output = df_output.append(additional_row4, ignore_index=True)
                            else:
                                if window.aircraft_type == 'Airbus':
                                    additional_row5['(R) Title'] = str(row['Cable_Type']) + cable_gauge + 'SH2'
                                    df_output = df_output.append(additional_row5, ignore_index=True)
                                else:
                                    df_output = df_output.append(additional_row5, ignore_index=True)
                    elif wire_type == 'TRIPLE SHIELD':
                        for k in range(3):
                            if k == 0:
                                if window.aircraft_type == 'Airbus':
                                    additional_row4['(R) Title'] = str(row['Cable_Type']) + str(
                                        row['Cable_Gauge']) + 'SH1'
                                    df_output = df_output.append(additional_row4, ignore_index=True)
                                else:
                                    df_output = df_output.append(additional_row4, ignore_index=True)
                            elif k == 1:
                                if window.aircraft_type == 'Airbus':
                                    additional_row5['(R) Title'] = str(row['Cable_Type']) + str(
                                        row['Cable_Gauge']) + 'SH2'
                                    df_output = df_output.append(additional_row5, ignore_index=True)
                                else:
                                    df_output = df_output.append(additional_row5, ignore_index=True)
                            else:
                                if window.aircraft_type == 'Airbus':
                                    additional_row6['(R) Title'] = str(row['Cable_Type']) + str(
                                        row['Cable_Gauge']) + 'SH3'
                                    df_output = df_output.append(additional_row6, ignore_index=True)
                                else:
                                    df_output = df_output.append(additional_row6, ignore_index=True)

        # Remove duplicate lines where the first two columns have identical values
        # df_output = df_output.drop_duplicates(subset=[df_output.columns[1], df_output.columns[10]], keep='first', ignore_index=True)

        # Save the output dataframe to a new XLSX file
    
        df_output = remove_slash_C(df_output)
        df_output.reset_index(drop=True)
        # remove_duplicates(df_output)
        df_output.to_excel(output_file, index=False)

        self.progress_bar.destroy()
        messagebox.showinfo("File Selected", 'File has been created successfully')
        #self.progress_bar.place(relx=0.485, rely=0.75, anchor=tk.CENTER)
        
        
        



     #def update_progress(window_list):
        #while len(window_list) == 0:
            #time.sleep(0.1)
        #window2 = window_list[0]
        #FileManager.update_progress()  
    


def function():
    global output_file
    global file_manager
    # Prompt the user to select the input XLSX file
    input_file = filedialog.askopenfilename(title="Select Input File",
                                            filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")))
    if not input_file:
        exit()
    # Ask user to select the output directory
    output_directory = askdirectory(title="Select Output Directory")
  
    # Define the output file path
    if window.aircraft_type == 'Airbus':
        output_file = output_directory + '/SeeEED_Airbus.xlsx'
    else:
        output_file = output_directory + '/SeeEED_Dassault.xlsx'

    # Read the input XLSX file
    file_manager = FileManager(input_file, output_file, window)
    file_manager.creating_progress_bar()
    thread = threading.Thread(target=file_manager.process_file)
    thread.start()

# Create a Tkinter window
window = tk.Tk()
window.title("SeeEED to 3DX Layout Handler")
window.iconbitmap(f"{current_directory}\\favicon.ico")
window.geometry("400x300")
window.resizable(width=False, height=False)
# Create a themed style for the window
style = ThemedStyle(window)
style.set_theme("elegance")  # Set the theme to "Arc"

# Set the window icon


# Load and set the background image

bg_image = Image.open(f"{current_directory}\\background aeros.jpg")
bg_photo = ImageTk.PhotoImage(bg_image)
bg_width = bg_photo.width()
bg_height = bg_photo.height() 
background_label = tk.Label(window, image=bg_photo)
background_label.place(x=0, y=0, relwidth=1, relheight=1)

# Welcome label
welcome_label = tk.Label(window, text="Welcome to the Layout Handler!", bg='#225491', fg='white')
welcome_label.pack(pady=0.1)

# Create radio buttons for aircraft selection
aircraft_var = tk.StringVar()
aircraft_var.set("Airbus")  # Default selection
tk.Radiobutton(window, text="Airbus", variable=aircraft_var, value=1, command=select_airbus, bg='#225d99', fg='white',
               activebackground='#225d99', activeforeground='white',
               selectcolor='#225d99').pack(anchor='w')
tk.Radiobutton(window, text="Dassault", variable=aircraft_var, value=2, command=select_dassault,
               bg='#225d99', fg='white', activebackground='#225d99', activeforeground='white',
               selectcolor='#225d99').pack(anchor='w')


# OK Button

ok_button = ttk.Button(window, text="OK", command=function)
ok_button.place(relx=0.485, rely=0.75, anchor=tk.CENTER)



# Run the Tkinter event loop
window.mainloop()



# file_manager.process_file()
