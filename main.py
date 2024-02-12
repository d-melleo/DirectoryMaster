# Imports
from tkinter import *
import tkinter as tk
from tkinter import ttk
from TkinterDnD2 import DND_FILES, TkinterDnD
from tkinter import filedialog as fd
from tkinter import messagebox
import pandas as pd
import csv
import string
import re
import os
import os.path
import sys
from typing import List, Set
from abc import ABC, abstractmethod
from PIL import Image, ImageTk
import threading
import random
import itertools
import time
try: 
    from gifs.gifs_list import gifs_list
except ImportError: gifs_list = []



root = TkinterDnD.Tk()



# Abstract class. Common structure for both windows.
class MyGui(ABC):
    def __init__(self) -> None:
        self.master: Tk = None
        self.color_primary: str = ""
        self.color_secondary: str = ""
        self.win_size_in_memory: str = None
        self.dataFrame: LabelFrame = None
        self.btn_Clear: Button = None
        self.viewBox: ttk.Treeview = None
        self.viewBox_label: Label = None
        self.buttonsFrame: Frame = None
        self.invalid_data: pd.DataFrame = None
        self.path_to_sourse_file: str = None
        self.path_to_output_file: str = None
        self.path_without_file_name: str = None
        self.column_id: int = 0
        # Pandas
        self.df: pd.DataFrame = None
        self.df_copy: pd.DataFrame = None
        # # GIF
        self.gif_path = ""
        self.gif_label: Label = None
        self.gif_count = 0
        self.gif_paths_list: List[str] = gifs_list
        self.gif_loop: str = None
        # Regex
        self.pattern = r"((?!1)(?!0)[0-9]{3}[\ \-\_\(\.))]*[\d]{3}[\ \-\_\(\)\.]*[\d$]{4})"


    def resource_path(self, relative_path):
        """ Get absolute path to resource, works for dev and for PyInstaller """
        try:
            # PyInstaller creates a temp folder and stores path in _MEIPASS
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")

        return os.path.join(base_path, relative_path)


    @abstractmethod # Read window size and position from config file.
    def set_geometry(self, conf: str) -> None:
        try:
            with open(conf, "r") as conf: 
                self.master.geometry(conf.read())
        except:
            pass

    @abstractmethod # Write window position and size upon closing to config file.
    def on_close(self, conf: str) -> None:
        try:
            with open(conf, "w") as conf: 
                conf.write(self.master.geometry())
        except: pass
        self.master.destroy()

    @abstractmethod # Configure GUI.
    def configureGui(self) -> None:
        # Root Window Parameters
        self.master.title('Directory Master')
        # try: self.master.iconbitmap('./master_roshi.ico')
        try: self.master.iconbitmap(self.resource_path("./master_roshi.ico"))
        except: pass
        self.master.minsize(320, 400)
        self.master.geometry("800x500")
        self.master.columnconfigure(1, weight=1)
        self.master.rowconfigure(0, weight=1)
        self.master.config(bg=self.color_primary)
        # Frame for Data Table
        self.dataFrame = LabelFrame(self.master, padx=0, pady=0)
        self.dataFrame.grid(row=0, column=1, rowspan=3, padx=0, pady=0, sticky=E+W+N+S)
        self.dataFrame.config(bg=self.color_primary)
        self.dataFrame.rowconfigure(0, weight=1)
        self.dataFrame.columnconfigure(0, weight=1)
        
        
        
        # Treeview style
        style = ttk.Style()
        style.configure("mystyle.Treeview", highlightthickness=1, bd=1, font=('Calibri', 10)) # Modify the font of the body
        style.configure("mystyle.Treeview.Heading", font=('Calibri', 11,'bold')) # Modify the font of the headings
        # style.layout("mystyle.Treeview", [('mystyle.Treeview.treearea', {'sticky': 'nswe'})]) # Remove the borders
        def fixed_map(option):
            return [elm for elm in style.map('Treeview', query_opt=option) if
                    elm[:2] != ('!disabled', '!selected')]
        style.map('Treeview', foreground=fixed_map('foreground'), background=fixed_map('background'))


        # Treeview widget for Data Table
        self.viewBox = ttk.Treeview(self.dataFrame, style="mystyle.Treeview", selectmode='none')
        self.viewBox.grid(row=0, column=0, pady=(50, 10), sticky=E+W+N+S)
        self.viewBox['show'] = ""
        self.viewBox.bind("<Button-1>", self.identify_column)
        treescrolly = tk.Scrollbar(self.dataFrame, orient="vertical", command=self.viewBox.yview)
        treescrollx = tk.Scrollbar(self.dataFrame, orient="horizontal", command=self.viewBox.xview)
        self.viewBox.configure(xscrollcommand=treescrollx.set, yscrollcommand=treescrolly.set)
        treescrolly.grid(row=0, column=3, rowspan=3, sticky=N+S, pady=(50,0))
        treescrollx.grid(row=0, column=0, columnspan=3, sticky=S+W+E)
        # Title for the Treeview Widget
        self.viewBox_label = Label(self.dataFrame, text="Select file", bg=self.color_primary)
        self.viewBox_label.grid(row=0, column=0, sticky=N+W+E)
        # Frame for Buttons
        self.buttonsFrame = Frame(self.master)
        self.buttonsFrame.grid(row=0, column=0, rowspan=10, sticky=N+S)
        self.buttonsFrame.config(bg=self.color_primary)
        
        self.btn_Clear_canvas = Canvas(self.dataFrame)
        self.btn_Clear_canvas.place(height=45, width=50, relx=1, x=-50)
        self.btn_Clear_canvas.config(highlightbackground=self.color_primary, bg=self.color_primary)
        
        self.configure_search(title="Search in (select column)", active=False)

        self.filter_gifts()
        if self.gif_paths_list != []:
            self.configure_gif_widget()

    # GIF
    def filter_gifts(self) -> List[str]:
        for i in self.gif_paths_list:
            try:
                Image.open(self.resource_path(i))
            except:
                self.gif_paths_list.remove(i)

    def configure_gif_widget(self):
        self.gif_path = random.choice(self.gif_paths_list)
        self.gif_path = self.resource_path(self.gif_path)
        self.gif_label = Label(self.buttonsFrame)
        self.gif_label.place(rely=1, relx=.5, relwidth=.9, anchor=CENTER, y=-50)
        self.gif_label.bind("<Button-1>", lambda x: self.run_anim(change=True))
        self.gif_label.config(bg=self.color_primary)
        self.gif_info = Image.open(self.gif_path)
        self.gif_frames = self.gif_info.n_frames
        self.gif_im = [PhotoImage(file=self.gif_path,format=f"gif -index {i}") for i in range(self.gif_frames)]

    def animation(self):
        self.gif_im2 = self.gif_im[self.gif_count]
        self.gif_label.configure(image=self.gif_im2)
        self.gif_count += 1
        if self.gif_count == self.gif_frames:
            self.gif_count = 0
        self.gif_loop = self.master.after(100, self.animation)

    @abstractmethod
    def run_anim(self, change:bool=False):
        if self.gif_paths_list != []:
            if change == True:
                self.master.after_cancel(self.gif_loop)
                self.gif_label.destroy()
                self.gif_count = 0
                self.configure_gif_widget()
            gif = threading.Thread(target=self.animation, name="gif")
            gif.start()
        else: return


    def configure_search(self, title: str, active: bool) -> Widget:
        for child in self.viewBox_label.winfo_children():
            child.destroy()
        # Search Widget Lable
        self.search_label = Label(self.viewBox_label, text='', bg=self.color_primary)
        self.search_label.config(text=title)
        self.search_label.pack(anchor=N+W)
        # Input box for Search Widget
        self.search_entry_var = StringVar()
        self.search_input_field = Entry(self.viewBox_label, textvariable=self.search_entry_var)
        self.search_input_field.pack(anchor=W)
        self.search_input_field['state'] = "disabled"
        self.search_entry_var.trace("w", self.filterTreeView)
        if active == True:
            self.search_input_field['state'] = "normal"
            self.search_input_field.focus_set()

    def identify_column(self, event):
        region = self.viewBox.identify("region", event.x, event.y)
        if region == "heading":
            column = self.viewBox.identify("column", event.x, event.y)
            self.column_id = int(column[1:]) -1
            search_label_title = "Search in column #" + str(self.column_id + 1)
            
            self.configure_search(search_label_title, active=True)

    def re_order_treeview(self, item, tag):
        search_var = self.viewBox.item(item)['values']
        self.viewBox.delete(item)
        self.viewBox.insert("", 0, values = search_var, tags=tag)

    def filterTreeView(self, *args):
        data = self.viewBox.get_children()
        search = self.search_entry_var.get().lower()
        search = str(search)
        
        for item in data:
            try:
                tag = self.viewBox.item(item)["tags"][0]
            except IndexError: pass
            
            str_item = self.viewBox.item(item)['values'][int(self.column_id)]
            str_item = str(str_item).lower()
            digit_item = re.sub(r"[\D\ +]", '', str(str_item))

            if digit_item == '':
                if search in str_item:
                    self.re_order_treeview(item, tag)
            
            elif digit_item != '':
                try:
                    if digit_item[0] == "1":
                        digit_item = digit_item[1:]
                except: pass
                try:
                    if search[0] == "1":
                        search = search[1:]
                except: pass
                if digit_item.startswith(search):
                    self.re_order_treeview(item, tag)


    # Clear data table from GUI.
    def clear_data(self) -> None:
        self.viewBox_label['text'] = "Select file"
        self.viewBox["show"] = ""
        self.viewBox.delete(*self.viewBox.get_children())
        self.configure_search(title="Search in (select column)", active=False)


    @abstractmethod # Write data table onto GUI.
    def write_on_screen(self, data: pd.DataFrame, tag: str, box_label: str = "Select file", clear:bool=True) -> None:
        # Clear the table
        if clear == True:
            self.clear_data()
        # Set a title for the window
        self.viewBox_label['text'] = box_label
        # Get cols
        # self.viewBox.grid(pady=(50, 10)) # Bring down a row for headers
        self.viewBox["columns"] = list(data.columns)
        self.viewBox["show"] = "headings"
        # Write data
        for column in self.viewBox["columns"]:
            self.viewBox.heading(column, text=column) # let the column heading = column name
        data_rows = data.to_numpy().tolist() # turns the dataframe into a list of lists
        for row in data_rows:
            self.viewBox.insert("", "end", values=row, tags=(tag,)) # inserts each list into the treeview.

    @abstractmethod # Partial save file function, common for both windows. Extended more into each class.
    def save_file(self) -> None:
        # Bring up filedialog window to choose a directory to save a file
        self.path_to_output_file: str = fd.asksaveasfilename()
        # If no path - return
        if not self.path_to_output_file: return None
        # Path to a file, without a name of a file
        get_path_without_file_name = self.path_to_output_file.split('/')[::-1][1:]
        get_path_without_file_name = get_path_without_file_name[::-1]
        self.path_without_file_name = '/'.join([str(elem) for elem in get_path_without_file_name])

    @abstractmethod # Both windows have even handlers for buttons.
    def eventsHandler(self, *args):
        pass


## Initial Window.
class MainWindow(MyGui):
    def __init__(self) -> None:
        super().__init__() # Get init objects from the abstract class MyGui.
        self.master: Tk = root # Set master as main root window
        self.color_primary: str = 'ivory2'
        self.color_secondary: str = 'azure'
        self.win_size_in_memory: str = "C:/Users/Public/main_win_conf.conf" # Config file with window parameters
        # Buttons
        self.btn_Browse: Button = None
        self.btn_Format: Button = None
        self.btn_Save: Button = None
        self.btn_RejectedData: Button = None
        self.btn_Close: Button = None
        # Saver variables
        self.output_file_length: int = None
        self.dnd_canvas: Canvas = None
        # Note indexes of invalid data for futher marking
        self.index_duplicated: Set(int) = set()
        self.index_missing_one: Set(int) = set()
        self.index_invalid_number: Set(int) = set()
        self.nan_values_in_num_indexes: Set[int] = set()

    # Set window position and size written from main_win_conf.conf
    def set_geometry(self) -> None:
        return super().set_geometry(self.win_size_in_memory)

    # Write window position and size into main_win_conf.conf upon closing.
    def on_close(self) -> None:
        super().on_close(self.win_size_in_memory)

    # Creat GUI objects.
    def configureGui(self) -> None:
        super().configureGui() # Use abstract class to configure GUI structure.
        self.set_geometry() # Set win pos and size
        # Write initial text with instructions on screen
        
        # Trigger eventsHandler func to manage buttons states
        self.master.bind("<Configure>", self.eventsHandler) # Will read buttons upon launching the App.
        self.master.bind("<ButtonRelease>", self.eventsHandler) # Will read buttons after each mouse-1 released click.
        # Register window closure and run on_close function.
        self.master.protocol("WM_DELETE_WINDOW",  self.on_close)
        self.viewBox.drop_target_register(DND_FILES)
        self.viewBox.dnd_bind("<<Drop>>", self.read_file)
        # Buttons
        self.btn_Browse = Button(self.buttonsFrame, text='Browse', width=10, command=lambda: self.browse_files())
        self.btn_Browse.grid(row=0, column=0, padx=(10), pady=7)
        self.btn_Browse.config(state="normal", bg=self.color_primary)
        self.btn_Browse.configure(highlightbackground=self.color_primary)

        self.btn_Format = Button(self.buttonsFrame, text='Format Data', width=10, command=lambda: self.format_data())
        self.btn_Format.grid(row=1, column=0, padx=(10), pady=7)
        self.btn_Format.config(state="disabled", bg=self.color_primary)
        self.btn_Format.configure(highlightbackground=self.color_primary)

        self.btn_Save = Button(self.buttonsFrame, text='Save', width=10, command=lambda: self.save_file())
        self.btn_Save.grid(row=2, column=0, padx=(10), pady=7)
        self.btn_Save.config(state="disabled", bg=self.color_primary)
        self.btn_Save.configure(highlightbackground=self.color_primary)

        self.btn_RejectedData = Button(self.buttonsFrame, text='View Rejected', width=10, command=lambda: self.launch_extraWindow())
        self.btn_RejectedData.grid(row=3, column=0, padx=(10), pady=7)
        self.btn_RejectedData.config(state="normal", bg=self.color_primary)
        self.btn_RejectedData.configure(highlightbackground=self.color_primary)

        self.btn_Close = Button(self.buttonsFrame, text='Close', width=10, command=self.on_close)
        self.btn_Close.grid(row=4, column=0, padx=(10), pady=7)
        self.btn_Close.config(state="normal", bg=self.color_primary)
        self.btn_Close.configure(highlightbackground=self.color_primary)

        # Button CLear
        self.btn_Clear = Button(self.btn_Clear_canvas, text="Clear", font=("Arial", 8, "italic"), bd=0, bg=self.color_primary, command= lambda: self.clear_data(all=True))
        self.btn_Clear.pack(side=RIGHT, padx=(0, 20), pady=(20, 0))

        self.configure_drag_n_drop()


    def configure_drag_n_drop(self):
        # Drag & Drop Icon on start
        self.dnd_canvas = Canvas(self.viewBox, bg="white", width=120, height=120, highlightbackground=self.color_primary)
        self.dnd_canvas.place(relx=.5, rely=.5, anchor=CENTER)
        self.dnd_image = Image.open(self.resource_path("./dnd_image.png"))
        self.dnd_image = ImageTk.PhotoImage(self.dnd_image)
        self.dnd_label: Label = Label(self.dnd_canvas, image=self.dnd_image)
        self.dnd_label.place(relx=.5, rely=.5, anchor=CENTER)
        self.dnd_label_text = Label(self.dnd_canvas)
        self.dnd_label_text.place(relx=.5, rely=.88, anchor=CENTER)
        self.dnd_label_text.config(text="Drag & Drop", bg="white", font=("Arial", 9, "italic"))
        # self.extensions_label = Label(self.viewBox, bg="white", text=".XLSX .XLS .CSV", font=("Arial", 7, "italic"))
        # self.extensions_label.place(relx=.5, rely=.65, anchor=CENTER)
        self.extensions_label = Label(self.dnd_canvas, bg="white", text=".XLSX .XLS .CSV", font=("Arial", 7, "italic"))
        self.extensions_label.place(relx=.5, rely=.15, anchor=CENTER)
        # Bind
        self.dnd_canvas.bind("<Button-1>", lambda x: self.browse_files())
        self.dnd_label.bind("<Button-1>", lambda x: self.browse_files())
        self.dnd_label_text.bind("<Button-1>", lambda x: self.browse_files())
        self.extensions_label.bind("<Button-1>", lambda x: self.browse_files())

    def forget_drag_n_drop(self):
        self.dnd_canvas.place_forget()
        self.dnd_label.place_forget()
        self.dnd_label_text.place_forget()
        self.extensions_label.place_forget()

    def run_anim(self, change:bool=False):
        super().run_anim(change)

    # Manage button states based on written data in GUI.
    def eventsHandler(self, *args) -> None:
        # Instructions
        if self.viewBox.tag_has(""):
            self.btn_Format['state'] = "disabled"
            self.btn_RejectedData.grid_forget() # Hides View Rejected button
            self.btn_Save['state'] = "disabled"
        # Unformatted data
        elif self.viewBox.tag_has("UnformattedData"):
            self.btn_Format['state'] = "normal"
            self.btn_RejectedData.grid_forget() # Hides View Rejected button
            self.btn_Save['state'] = "disabled"
            self.dnd_canvas.place_forget()
            self.extensions_label.place_forget()
        # Formatted data
        elif self.viewBox.tag_has("FormattedData"):
            self.btn_Format['state'] = "disabled"
            self.btn_Save['state'] = "normal"
            self.btn_RejectedData.grid(row=3, column=0, padx=(10), pady=7) # Creates button View Rejected.
            self.dnd_canvas.place_forget()
            self.extensions_label.place_forget()
        else:
            try:
                len(self.df) # If df has no data Format button will be disabled.
            except:
                self.btn_Format['state'] = "disabled"
                self.btn_Save['state'] = "disabled"
            try:
                len(self.invalid_data) # Creates button View Rejected if rejected data presents.
                self.btn_RejectedData.grid(row=3, column=0, padx=(10), pady=7)
            except:
                self.btn_RejectedData.grid_forget() # Hides View Rejected button


    # Trigerred by Browse button
    def browse_files(self) -> str:
        # Filedialog window allows to choose .csv or excel file
        self.path_to_sourse_file = fd.askopenfilename(title="Select file", filetypes=(("", ".xlsx .xls .csv"),("CSV", "*.csv"), ("EXCEL", ".xlsx .xls")))
        if not self.path_to_sourse_file: return # If file not selected will return
        else: self.read_file()

    # Write data to pd.DataFrame based on extension.
    def read_file(self, event=None) -> pd.DataFrame:
        try:
            if event.data.endswith((".xlsx", ".xls", ".csv")):
                self.path_to_sourse_file = event.data
        except: pass
        # Read a file depending on extension. And bring pop-up windown with error messages when needed.
        try:
            # CSV files
            if self.path_to_sourse_file.split('.')[::-1][0] == 'csv':
                self.df = pd.read_csv(self.path_to_sourse_file)
            # XLSX files
            elif self.path_to_sourse_file.split('.')[::-1][0] in ['xlsx', 'xls']:
                self.df = pd.read_excel(self.path_to_sourse_file, dtype=str)
            
            # Write data table on the screen
            tag = "UnformattedData"
            box_label = "{} \tRows: {}".format(self.path_to_sourse_file, len(self.df))
            self.write_on_screen(self.df, tag, box_label) # Write data onto GUI window
            
        # Errors handling
        # If can't open a file due to 'error13: access denied'.
        except PermissionError:
            messagebox.showinfo('Can\'t open a file', f'File can\'t be opened at \n{self.path_to_sourse_file} \nMove this file to \'C:/Users/Public\' and try again.')
        # If chosen file has invalid extension.
        except TypeError:
            messagebox.showinfo("Information", "Can\'t read this extension.")
        # If file doesn't exist
        except FileNotFoundError:
            messagebox.showinfo("Information", f"File not found: \n{self.path_to_sourse_file}")
        # If empty file.
        except ValueError:
            messagebox.showinfo("Information", f"This file is empty.")
        except OSError:
            messagebox.showinfo("Information", f"Choose one file")


    def clear_data(self, all: bool = False) -> None:
        super().clear_data()
        if all == True:
            self.forget_drag_n_drop()
            self.configure_drag_n_drop()
            self.df = None
            self.df_copy = None
            self.invalid_data = None
            self.clear_prior_data_indexes()

    def clear_prior_data_indexes(self):
        self.index_duplicated: Set(int) = set()
        self.index_missing_one: Set(int) = set()
        self.index_invalid_number: Set(int) = set()
        self.nan_values_in_num_indexes: Set[int] = set()


    # Writes data frame on GUI window.
    def write_on_screen(self, data: pd.DataFrame, tag: str, box_label: str = "Select file") -> None:
        return super().write_on_screen(data, tag, box_label)


    def note_indexes_for_invalid_data(self, sort: str, col: str=None):
        if col != None:
            df: pd.DataFrame = self.df[col].dropna()
        else:
            df: pd.DataFrame = self.df.dropna(axis=0, how='any')

        if sort == "missing one":
            indexes = [x for x in self.df_copy.index if x not in df.index and x not in self.index_invalid_number and x not in self.index_duplicated]
            self.index_missing_one = self.index_missing_one.union(set(indexes))
            
        elif sort == "invalid number":
            indexes = [x for x in self.df_copy.index if x not in df.index and x not in self.index_missing_one and x not in self.index_duplicated]
            self.index_invalid_number = self.index_invalid_number.union(set(indexes))
            
        elif sort == "duplicates":
            indexes = [x for x in self.df_copy.index if x not in df.index and x not in self.index_invalid_number and x not in self.index_missing_one]
            self.index_duplicated = self.index_duplicated.union(set(indexes))
            


    # Write columns as first row.
    def move_cols_down_to_row(self, df: pd.DataFrame) -> pd.DataFrame:
        df.loc[-1] = df.columns # Write columns into first row in index -1
        df.index = df.index + 1 # Correct indexes to start with 0
        df = df.sort_index() # Sort indexes
        return df

    # If data has one column it will be split into two columns by applying regex to identify names and numbers.
    def split_if_single_column(self, df: pd.DataFrame) -> pd.DataFrame:
        df["Name"] = df[df.columns[0]].apply(lambda x: re.sub(self.pattern, '', str(x))) # Get name by removing numbers from column
        df["Number"] = df[df.columns[0]].apply(lambda x: re.findall(self.pattern, str(x))) # Pull number by applying regex.
        df["Number"] = df["Number"].apply(lambda x: x if re.findall(self.pattern, str(x)) else pd.np.nan) # Remove empty lists from row
        df = df[["Name", "Number"]] # Make data frame with two columns.
        return df

    # Looking for column with the most data if ambiguous columns found
    def get_name_for_column_with_more_data(self, df: pd.DataFrame, tag: str) -> str:
        df = df.replace('\W',' ', regex=True) # Remove all special characters
        # Get column name with the most rows
        last_length = 0
        column_name = ''
        for column in df.columns:
            if tag == "name": # Remove rows with plain digits (doesn't remove if at least one letter)
                df[column] = df[column].replace({r"^[0-9\ +]+$": pd.np.nan}, regex=True)
            elif tag == "number": # Apply regex to keep only numbers in rows
                nan_values = df[df[column].isna()]
                df[column] = df[column].apply(lambda x: re.findall(self.pattern, str(x))) # Find all numbers and will output them as list
                df[column] = df[column].apply(lambda x: x if re.findall(self.pattern, str(x)) else pd.np.nan) # Remove empty lists from row
            temp_length = df[column].notnull().sum() # Remove empty rows from col and get col length.
            # Compare length to last column and return the longest column
            if temp_length >= last_length:
                last_length = temp_length
                column_name = column
                try:
                    self.nan_values_in_num_indexes = set(nan_values.index.values)
                except: pass
        # Rename column in case if duplicates in column names
        renamed_col = str(column_name) + str(tag)
        # Write formatted column into df
        try:
            self.df[renamed_col] = df[column_name]
        except KeyError:
            pass
        # Return name of a column
        return renamed_col

    # Find columns which contain keywords: name, phone, number, cell, mobile.
    def return_name_for_column_with_more_data(self, df: pd.DataFrame, lookup_list: List[str], tag: str) -> str:
        # Search for columns that contain keywords in lookup_list
        try_df = df[[x for x in df.columns if any(y for y in lookup_list if y in str(x).lower()) and "unnamed" not in str(x).lower()]]
        if not try_df.empty: # If found columns based on the lookup_list - leave only one col with the most entries
            col = self.get_name_for_column_with_more_data(try_df, tag)
        else: # If no cols titled as in lookup_list - check all cols and locate the correct one.
            col = self.get_name_for_column_with_more_data(df, tag)
        # Return a name of the column
        return col

    # Remove empty rows
    def filter_empty_rows(self, df: pd.DataFrame) -> pd.DataFrame:
        for column in df.columns:
            df[column] = df[column].apply(lambda x: pd.np.nan if x == "" else x)
            df = df[df[column].notna()]
            df = df[df[column].notnull()]
        return df

    # Remove special characters and anything which is not in [0-9] from numbers
    def remove_special_char_from_numbers(self, number: pd.DataFrame) -> pd.Series:
        number_ = str(number["Number"])
        return re.sub(r'[^0-9]+', '', number_)

    # Remove white spaces and trim all numbers to 10 digits
    def filter_numbers(self, col: pd.DataFrame) -> pd.Series:
        col_ = str(col["Number"]).translate(str.maketrans('', '', string.punctuation))
        col_ = col_.replace(r"\ +", '')
        if len(col_) == 10 and (col_)[0] != '1':
            return col_
        elif len(col_) >= 11 and col_[0] == '1':
            return col_[1:11]
        else:
            return pd.np.nan

    ## Main Logic. Trigerred by Format button
    def format_data(self) -> pd.DataFrame:
        self.clear_prior_data_indexes()
        # Bring down columns to row 1 in case if columns are not titled.
        self.df = self.move_cols_down_to_row(self.df)
        # Make a copy of the original file in memory to get rejected data later on.
        self.df_copy = self.df.copy()
        # Split a single column in file.
        if len(self.df.columns) == "1":
            self.df = self.split_if_single_column(self.df)
            # Get indexes of rows with NaN cells
            self.note_indexes_for_invalid_data(sort="missing one")
        # Get names column
        names_col = self.return_name_for_column_with_more_data(self.df, ["name"], tag="name")
        # Get indexes of rows with NaN cells in names
        self.note_indexes_for_invalid_data(sort="missing one", col=names_col)
        # Get numbers column
        numbers_col = self.return_name_for_column_with_more_data(self.df, ["phone", "number", "cell", "mobile"], tag="number")
        # Get indexes of rows with invalid numbers
        self.index_missing_one = self.index_missing_one.union(self.nan_values_in_num_indexes)
        self.note_indexes_for_invalid_data(sort="invalid number", col=numbers_col)
        # Make data frame with found above columns
        try:
            self.df = self.df[[names_col, numbers_col]]
        except KeyError:
            return

        # Rename columns
        self.df.rename(columns={
            names_col: "Name",
            numbers_col: "Number"
        }, inplace=True)
        
        # Let's filter out too long values
        self.df["Number"] = self.df.apply(lambda x: self.remove_special_char_from_numbers(x), axis=1)
        self.df["Number"] = self.df.apply(lambda x: self.filter_numbers(x), axis=1)
        # Get indexes of rows with invalid numbers
        self.note_indexes_for_invalid_data(sort="invalid number", col="Number")
        # Remove empty rows and special characters
        self.df = self.filter_empty_rows(self.df)
        # Get indexes of rows with NaN cells
        self.note_indexes_for_invalid_data(sort="missing one")
        # Remove non word characters from Names
        self.df = self.df.replace('\W',' ', regex=True)
        # Remove whitespaces in Name column
        self.df["Name"] = self.df["Name"].str.strip()
        self.df["Name"] = self.df["Name"].apply(lambda x: re.sub(r"\ {2,}", ' ', x))
        # Sort & remove duplicates
        self.df = self.df.iloc[self.df["Name"].str.lower().argsort()]
        self.df.drop_duplicates(subset="Name", keep = False, inplace = True)
        self.df.drop_duplicates(subset="Number", keep = False, inplace = True)
        self.note_indexes_for_invalid_data(sort="duplicates")


        # Write data table onto GUI window
        tag = "FormattedData"
        box_label = "{} — Rows: {}".format(self.path_to_sourse_file, len(self.df))
        self.write_on_screen(self.df, tag, box_label)
        
        try:
            # Get rejected data
            self.invalid_data = self.get_rejected_data()
            # Ask if show rejected data
            self.show_rejected_data()
        except: pass



    # Get rejected data.
    def get_rejected_data(self) -> pd.DataFrame:
        # Find removed indexes from a copy of the original file
        indexes = [x for x in self.df_copy.index if x not in self.df.index]
        # Write removed indexes as rejected data
        invalid_data = self.df_copy.iloc[indexes]
        # See self.move_cols_down_to_row. Columns become as row with index 0.
        # If index 0 not in rejected data, it means it was taken as a valid contact. Therefore it's not a header.
        # Then columns in rejected data need to be unnamed.
        if 0 not in indexes:
            invalid_data.columns = ['' for x in invalid_data.columns]
        elif 0 in indexes:
            cols = invalid_data.columns # Get columns
            first_row = invalid_data.iloc[0] # Get first row
            col_info = cols.difference(first_row) # Compare columns to the first row if are the same
            if col_info.empty:
                # If columns are the same as first row - remove first row
                invalid_data = invalid_data.drop([0])
        return invalid_data

    # Message box. Ask if view rejected data now.
    def show_rejected_data(self) -> None:
        length = len(self.invalid_data)
        if length >= 1:
            question = messagebox.askquestion("Question", f"Rejected {length} rows. Would you like to view rejected data?")
            if question == 'yes':
                self.launch_extraWindow() # Launch second GUI window with rejected data.
            else: pass

    # Run second GUI window to show rejected data.
    def launch_extraWindow(self):
        gui = ExtraWindow(
            self.invalid_data,
            self.index_missing_one,
            self.index_invalid_number,
            self.index_duplicated,
        )
        gui.configureGui()
        gui.run_anim()
        gui.prepare_data_before_writting()
        gui.put_all_data_on_screen(show_all=True)

        
    # Save files by 500 entries.
    def csv_saver(self, rows: int, path: str, file_name: str, file_count: int) -> None:
        if os.path.exists(path): # Check if file name exists to avoid overwriting.
            # Ask if overwrite
            overwrite_first_file = messagebox.askquestion('Warning!', f'File {file_name} already exists. Would you like to overwrite?', icon="warning")
            if overwrite_first_file == 'yes': 
                self.df[rows:rows+500].to_csv(path, index=False, header=True, quoting=csv.QUOTE_NONNUMERIC) # Write data as .csv by 500 rows each
                file_count += 1 # File count needed to auto name files
                self.output_file_length -= 500 # Keep track of saved rows
            else:
                return
        else: # If file with same name doesn't exist, then save file as .csv by 500 rows each.
            self.df[rows:rows+500].to_csv(path, index=False, header=True, quoting=csv.QUOTE_NONNUMERIC)
            file_count += 1 # File count needed to auto name files
            self.output_file_length -= 500 # Keep track of saved rows

    # File saving logic. Triggered by Save button. 
    def save_file(self) -> None:
        super().save_file() # Open file dialog, choose path to save.
        if not self.path_to_output_file: return # If no path selected return.
        file_count = 1 # File count needed to auto name files
        self.output_file_length = len(self.df) # Keep track of saved rows
        ## Will write all formatted data as .csv files. Each file has no more, than 500 entries.
        try:
            if len(self.df) >= 1: # Check if has data to save.
                for rows in range(len(self.df)):
                    first_file_name = self.path_to_output_file.split('/')[::-1][0] + '.csv'# Name first file as user decided
                    # All fies after first one will be named automatically based on first file name + (1), (2) and so on
                    other_file_names = self.path_to_output_file.split('/')[::-1][0] + "(" + str(file_count) + ")" + '.csv'
                    if rows % 500 == 0: # Make sure saved file has 500 rows or less
                        if file_count == 1:
                            self.csv_saver(rows=rows, path=str(self.path_to_output_file) + '.csv', file_name=first_file_name, file_count=file_count)
                            file_count += 1 # File count needed to auto name files
                        else:
                            self.csv_saver(rows=rows, path=str(self.path_to_output_file) + "(" + str(file_count) + ")" + '.csv', file_name=other_file_names, file_count=file_count)
                            file_count += 1 # File count needed to auto name files
            else: # No data to save.
                messagebox.showwarning('Can\'t save a file', 'File has no data to write')
                return # Return back to Main GUI window
        except PermissionError: # An error may occur.
            messagebox.showinfo('Can\'t save a file', f'File can\'t be saved at \n{self.path_without_file_name} \nas you don\'t have admin privileges.\nTry to save this file at \'C:/Users/Public\'.', icon="error")
            return # Return back to Main GUI window


# Second GUI window for rejected data.
class ExtraWindow(MyGui):
    def __init__(
        self,
        invalid_data,
        index_missing_one: Set[int],
        index_invalid_number: Set[int],
        index_duplicated: Set[int]
    ) -> None:
        super().__init__() # Get init objects from the abstract class MyGui.
        self.master: Toplevel = Toplevel(root) # Set master as a child window of the Main GUI window.
        self.color_primary: str = 'azure'
        self.color_secondary: str = 'ivory2'
        self.win_size_in_memory: str = "C:/Users/Public/extra_win_conf.conf" # Config file with window parameters
        self.win_size_in_memory = self.resource_path(self.win_size_in_memory)
        self.invalid_data: pd.DataFrame = invalid_data # Passed rejected data from parent class MainWindow.
        self.saving_data: pd.DataFrame = None
        # Sorted indexes for invalid data
        self.index_missing_one = index_missing_one
        self.index_invalid_number = index_invalid_number
        self.index_duplicated = index_duplicated
        # Converted indexes to lists
        self.id_missing_cell: List[int] = None
        self.id_invalid_number: List[int] = None
        self.id_duplicated: List[int] = None
        self.id_unmarked: List[int] = None
        # Data frames with indexes separately
        self.missing_cell: pd.DataFrame = None
        self.invalid_number: pd.DataFrame = None
        self.duplicated: pd.DataFrame = None
        self.unmarked: pd.DataFrame = None
        # Buttons
        self.btn_Save: Button = None
        self.btn_Close: Button = None
        # Sort Labels
        self.btn_all_data: Label = None
        self.yellow_label: Label = None
        self.orange_label: Label = None
        self.blue_label: Label = None
        self.grey_label: Label = None

    # Set window position and size written from main_win_conf.conf
    def set_geometry(self) -> None:
        return super().set_geometry(self.win_size_in_memory)

    # Write window position and size into main_win_conf.conf upon closing.
    def on_close(self) -> None:
        return super().on_close(self.win_size_in_memory)

    # Creat GUI objects.
    def configureGui(self) -> None:
        super().configureGui() # Use abstract class to configure GUI structure.
        self.set_geometry() # Set win pos and size
        # Trigger eventsHandler func to manage buttons states
        self.master.title('Rejected Data')
        self.master.bind("<Configure>", self.eventsHandler)
        self.viewBox.bind("<ButtonRelease>", self.eventsHandler)
        # Register window closure and run on_close function.
        self.master.protocol("WM_DELETE_WINDOW",  self.on_close)
        # Canvas for notes
        info_canvas = Canvas(self.buttonsFrame, bd=0, highlightthickness=0, bg=self.color_primary, height=150)
        info_canvas.place(relwidth=.9, relx=.5, rely=1, y=-200, anchor=CENTER)
        # Marks
        self.btn_Clear_canvas.place_configure(width=100, x=-100)
        self.btn_all_data = Button(self.btn_Clear_canvas, text="show all", font=("Arial", 8, "italic"), bd=0, bg=self.color_primary, activebackground=self.color_primary, command=lambda: self.put_all_data_on_screen(show_all=True))
        self.btn_all_data.pack(side=RIGHT, padx=(0, 20), pady=(20, 0))
        self.yellow_label = Label(info_canvas, bd=0, text="Empty cell", bg="yellow", width=12, height=1, font=('Calibri', 10, 'italic'))
        self.yellow_label.grid(row=1, column=0, pady=3)
        self.orange_label = Label(info_canvas, bd=0, text="Invalid number", bg="orange", width=12, height=1, font=('Calibri', 10, 'italic'))
        self.orange_label.grid(row=2, column=0, pady=3)
        self.blue_label = Label(info_canvas, bd=0, text="Duplicate", bg="blue", width=12, height=1, font=('Calibri', 10, 'italic'))
        self.blue_label.grid(row=3, column=0, pady=3)
        self.grey_label = Label(info_canvas, bd=0, text="Other", bg="grey", width=12, height=1, font=('Calibri', 10, 'italic'))
        self.grey_label.grid(row=4, column=0, pady=3)
        # Label binds
        self.yellow_label.bind("<Button-1>", lambda x: self.show_selected_data(self.missing_cell, "MissingCell", "missing data"))
        self.orange_label.bind("<Button-1>", lambda x: self.show_selected_data(self.invalid_number, "InvalidNumber", "invalid numbers"))
        self.blue_label.bind("<Button-1>", lambda x: self.show_selected_data(self.duplicated, "Duplicate", "duplicates"))
        self.grey_label.bind("<Button-1>", lambda x: self.show_selected_data(self.unmarked, "Unmarked", "other defects"))
        # Button to save invalid data as excel file
        self.btn_Save = Button(self.buttonsFrame, text='Save as excel', width=10, command=lambda: self.save_file(self.saving_data))
        self.btn_Save.grid(row=1, column=0, padx=(10), pady=7)
        self.btn_Save.config(bg=self.color_primary)
        self.btn_Save.configure(highlightbackground=self.color_primary)
        # Close button
        self.btn_Close = Button(self.buttonsFrame, text='Close', width=10, command=self.on_close)
        self.btn_Close.grid(row=4, column=0, padx=(10), pady=7)
        self.btn_Close.config(state="normal", bg=self.color_primary)
        self.btn_Close.configure(highlightbackground=self.color_primary)

    def run_anim(self, change:bool=False):
        super().run_anim(change)

    # Manage button states based on written data in GUI.
    def eventsHandler(self, *args) -> None:
        try:
            if len(self.invalid_data) > 0: # If no rejected data - disable Save button
                self.btn_Save.config(state="normal")
            else: self.btn_Save.config(state="disabled")
        except:
            self.btn_Save.config(state="disabled")

        tags = set()
        for child in self.viewBox.get_children():
            tags = tags.union(set(self.viewBox.item(child)["tags"]))
        self.update_saving_data(tags)

    def update_saving_data(self, tags: set):
        # Declare tags to data
        data_reference = {
            "MissingCell": self.id_missing_cell,
            "InvalidNumber": self.id_invalid_number,
            "Duplicate": self.id_duplicated,
            "Unmarked": self.id_unmarked
        }
        # Get indexes fro data currently viewed
        data_indexes = [data_reference.get(x) for x in data_reference.keys() if x in tags]
        data_indexes = list(itertools.chain(*data_indexes))
        data_indexes = sorted(data_indexes)
        self.saving_data = self.invalid_data[self.invalid_data.index.isin(data_indexes)]
    
    # Writes data frame on GUI window.
    def write_on_screen(self, data: pd.DataFrame, clear:bool=False, tag:str="RejectedData") -> None:
        super().write_on_screen(
            data = data,
            tag = tag,
            box_label = "Rejected {} rows".format(len(self.invalid_data)),
            clear = clear
        )


    def show_selected_data(self, data: pd.DataFrame, tag: str, reason: str):
        self.viewBox.delete(*self.viewBox.get_children())
        data_rows = data.to_numpy().tolist() # turns the dataframe into a list of lists
        for row in data_rows:
            self.viewBox.insert("", "end", values=row, tags=(tag,)) # inserts each list into the treeview.
        self.viewBox_label['text'] = "Rejected {} rows with {}".format(len(data), reason)
        
        self.resize_pressed_button(tag)


    def resize_pressed_button(self, tag:str=None) -> None:
        buttons_reference = {
            "MissingCell": self.yellow_label,
            "InvalidNumber": self.orange_label,
            "Duplicate": self.blue_label,
            "Unmarked": self.grey_label
        }
        
        if tag:
            btn = buttons_reference.get(tag)
            btn.config(height=2)
            for b in buttons_reference.values():
                if b != btn:
                    b.config(height=1)
            self.btn_all_data.config(font=("Arial", 9, "italic"), fg="green")
        else:
            for b in buttons_reference.values():
                b.config(height=1)
            self.btn_all_data.config(font=("Arial", 8, "italic"), fg='black')


    def prepare_data_before_writting(self):
        # Marked indexes with invalid data
        self.id_missing_cell = sorted(list(self.index_missing_one))
        self.id_invalid_number = sorted(list(self.index_invalid_number))
        self.id_duplicated = sorted(list(self.index_duplicated))
        # Anything that did not match marking criterias
        indexes = list(itertools.chain(self.id_missing_cell, self.id_invalid_number, self.id_duplicated))
        self.id_unmarked = [x for x in self.invalid_data.index if x not in indexes]
        # Separate dataframes based on indexes
        self.missing_cell = self.invalid_data[self.invalid_data.index.isin(self.id_missing_cell)]
        self.invalid_number = self.invalid_data[self.invalid_data.index.isin(self.id_invalid_number)]
        self.duplicated = self.invalid_data[self.invalid_data.index.isin(self.id_duplicated)]
        self.unmarked = self.invalid_data[self.invalid_data.index.isin(self.id_unmarked)]
    
    def put_all_data_on_screen(self, show_all:bool=False):
        if show_all == True:
            self.clear_data()
        # Write above data frames with tags
        self.write_on_screen(data=self.missing_cell, tag="MissingCell")
        self.write_on_screen(data=self.invalid_number, tag="InvalidNumber")
        self.write_on_screen(data=self.duplicated, tag="Duplicate")
        self.write_on_screen(data=self.unmarked, tag="Unmarked")
        # Set colors of rows for written data
        self.viewBox.tag_configure("MissingCell", background="yellow")
        self.viewBox.tag_configure("InvalidNumber", background="orange")
        self.viewBox.tag_configure("Duplicate", background="blue")
        self.viewBox.tag_configure("Unmarked", background="grey")
        
        self.resize_pressed_button()

    # File saving logic. Triggered by Save button.
    def save_file(self, data: pd.DataFrame) -> None:
        super().save_file() # Open file dialog, choose path to save.
        if not self.path_to_output_file:
            self.master.focus_set()
            return # If no path selected return.
        # Get name of a file, without showing a path to a file
        name_of_file = self.path_to_output_file.split('/')[::-1][0] + '.xlsx'
        try:
            if os.path.exists(self.path_to_output_file + '.xlsx'): # Check if file name exists to avoid overwriting.
                # Ask if overwrite
                overwrite_invalid_data_file = messagebox.askquestion('Warning', f'File {name_of_file} already exists. Would you like to overwrite?', icon="warning")
                if overwrite_invalid_data_file == 'yes':
                    data.to_excel(self.path_to_output_file + '.xlsx', index=False) # Save file as excel.
                else:
                    return # If not overwrite - return back to GUI
            else: # If file with same name doesn't exist, then save file 
                data.to_excel(self.path_to_output_file + '.xlsx', index=False) # Save file as excel.
        except PermissionError: # An error may occur.
            messagebox.showerror('Can\'t save a file', f'File can\'t be saved at \n{self.path_without_file_name} \nas you don\'t have admin privileges.\nTry to save this file at \'C:/Users/Public\'')
            return # Return back to GUI


# Run Main GUI
if __name__ == '__main__':
    gui = MainWindow()
    gui.configureGui()
    gui.run_anim()
    gui.master.mainloop()



