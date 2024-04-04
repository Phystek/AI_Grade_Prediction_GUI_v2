#using version 5.22 of customtkinter
import customtkinter as ctk
from settings import *
from tkinterdnd2 import TkinterDnD, DND_ALL
from tkinter import filedialog, Canvas, StringVar, ttk,BooleanVar, DISABLED
import pandas as pd
from tkinter import font as tkFont
import openpyxl
from CTkToolTip import *
from PIL import Image, ImageTk, ImageOps

#importing pyautogui will cause the app to not rescale when moving between windows, poor resolution, but stable columns!
import pyautogui
#from screeninfo import get_monitors
import ctypes

#load color theme
ctk.set_default_color_theme('phystek_colours.json')
#ctk.set_default_color_theme('dark-blue')



class MainApp(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self):

        # setup
        super().__init__()
        ctk.set_appearance_mode("dark")
        w, h = 800, 450  # Width and height.
        x, y = 200, 200  # Screen position.
        self.geometry('%dx%d+%d+%d' % (w, h, x, y))
        self.minsize(800,450)
        self.title('AI Grade Prediction')
        self.TkdndVersion = TkinterDnD._require(self)

        self.resizable(True, True)  # Disable resizing of the window
        #trigger resize function on window configuration change



        self.bind("<Configure>", self.on_resize)

        #canvas data (not currently used)
        self.canvas_width = 0
        self.canvas_height = 0
        btn_height = 26
        btn_width = 135
        self.check_vars = []


        #main frame here, holds all of the other frames with their widgets
        frame_main = ctk.CTkFrame(master=self, corner_radius=10, fg_color=MAIN_FRAME_COLOR, bg_color=MAIN_FRAME_COLOR)
        frame_main.pack(padx=0, pady=0, fill="both", expand=True)

        #add prediction and training tabs to the main frame
        self.main_tabs = ctk.CTkTabview(master=frame_main, corner_radius=10, fg_color=LABEL_COLOR,
                                               bg_color=MAIN_FRAME_COLOR)
        self.main_tabs._segmented_button.grid(sticky="W")

        self.main_tabs.pack(padx=5, pady=(0,5), ipadx=0, ipady=0, fill="both", expand=True, side="left")
        self.main_tabs.add("Train Model")
        self.main_tabs.add("Predict Grades")
        self.main_tabs._segmented_button.grid(sticky="W")

        #add frame to the training tab
        self.train_frame = ctk.CTkFrame(master=self.main_tabs.tab("Train Model"), corner_radius=10, fg_color=MAIN_FRAME_COLOR)
        self.train_frame.pack(padx=0, pady=0, fill="both", expand=True)
        self.train_frame.rowconfigure(0, weight=1)
        self.train_frame.columnconfigure(0, weight=1)
        self.train_frame.columnconfigure(1, weight=4)

        #make left and right frames for the gui
        self.train_frame_left = ctk.CTkFrame(master=self.train_frame, corner_radius=10, fg_color=LABEL_COLOR, width=200)
        self.train_frame_left.grid(row=0, column=0, padx=5, pady=5, sticky='NSEW')
        self.train_frame_left.rowconfigure(0, weight=1)
        self.train_frame_left.rowconfigure(1, weight=1)
        self.train_frame_left.rowconfigure(2, weight=10)
        self.train_frame_left.columnconfigure(0, weight=1)
        self.train_frame_left.grid_propagate(False)

        self.train_frame_right = ctk.CTkFrame(master=self.train_frame, corner_radius=10, fg_color=LABEL_COLOR, width=400)
        self.train_frame_right.grid(row=0, column=1, padx=5, pady=5, sticky='NSEW')
        self.train_frame_right.rowconfigure(0, weight=1)
        self.train_frame_right.rowconfigure(1, weight=1)
        self.train_frame_right.rowconfigure(2, weight=20)
        self.train_frame_right.rowconfigure(3, weight=1)
        #self.train_frame_right.rowconfigure(3, weight=1)
        self.train_frame_right.columnconfigure(0, weight=1)
        self.train_frame_right.grid_propagate(False)

        self.import_frame = ctk.CTkFrame(master=self.train_frame_left, corner_radius=10, fg_color=LABEL_COLOR)
        self.import_frame.grid(row=0, column=0, padx=5, pady=5, sticky='NEW')
        self.import_frame.rowconfigure(0, weight=1)
        self.import_frame.rowconfigure(1, weight=1)
        self.import_frame.columnconfigure(0, weight=1)




        #add widgets to left frame
        self.import_file_btn = (ctk.CTkButton(self.import_frame, text='Import Gradebook File', height=btn_height, command=self.file_dialog,
                                          font=("Inter", 14, "bold")))
        self.import_file_btn.grid(row=0, column=0, columnspan=1, padx=5, pady=10, sticky='NEW')
        self.filename_var = StringVar(value="Filename: No files loaded")
        self.entryWidget = ctk.CTkEntry(master=self.import_frame,font=("Inter", 10, "italic"), textvariable=self.filename_var, height=22)
        self.entryWidget.grid(row=1, column=0, columnspan=1, padx=5, pady=3, ipadx=0, ipady=0, sticky='NEW')

        self.entryWidget.drop_target_register(DND_ALL)
        self.entryWidget.dnd_bind("<<Drop>>", self.get_path)

        self.stored_data_frame = ctk.CTkFrame(master=self.train_frame_left, corner_radius=10, fg_color=LABEL_COLOR)
        self.stored_data_frame.grid(row=1, column=0, padx=5, pady=5, sticky='NEW')
        self.stored_data_frame.rowconfigure(0, weight=1)
        self.stored_data_frame.rowconfigure(1, weight=1)
        self.stored_data_frame.rowconfigure(2, weight=1)
        self.stored_data_frame.columnconfigure(0, weight=1)
        self.stored_data_frame.columnconfigure(1, weight=1)

        self.stored_files_label = ctk.CTkLabel(master=self.stored_data_frame, text="Number of files stored:", font=("Inter", 12))
        self.stored_files_label.grid(row=0, column=0, padx=5, pady=5, sticky='NW')
        self.no_of_datafiles_SV = StringVar(value="0")
        self.num_files_label = ctk.CTkLabel(master=self.stored_data_frame, textvariable=self.no_of_datafiles_SV, font=("Inter", 12, "bold"))
        self.num_files_label.grid(row=0, column=1, padx=1, pady=5, sticky='NW')

        #self.no_of_datafiles_SV.set("test")

        self.check_store_1 = BooleanVar(value=False)
        self.file_1_name = StringVar(value="No Files Stored")
        self.check_file_1 = ctk.CTkCheckBox(master=self.stored_data_frame, variable=self.check_store_1,state=DISABLED, textvariable=self.file_1_name, width=12, height=12, font=("Inter", 10))
        self.check_file_1.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky='NW')

        self.check_store_2 = BooleanVar(value=False)
        self.file_2_name = StringVar(value="No Files Stored")
        self.check_file_2 = ctk.CTkCheckBox(master=self.stored_data_frame, variable=self.check_store_2, state=DISABLED,
                                            textvariable=self.file_2_name, width=12, height=12, font=("Inter", 10))
        self.check_store_3 = BooleanVar(value=False)
        self.file_3_name = StringVar(value="No Files Stored")
        self.check_file_3 = ctk.CTkCheckBox(master=self.stored_data_frame, variable=self.check_store_3, state=DISABLED,
                                            textvariable=self.file_3_name, width=12, height=12, font=("Inter", 10))


        self.train_buttons_frame = ctk.CTkFrame(master=self.train_frame_left, corner_radius=10, fg_color=LABEL_COLOR)
        self.train_buttons_frame.grid(row=2, column=0, padx=5, pady=5, sticky='SEW')
        self.train_buttons_frame.rowconfigure(0, weight=1)
        self.train_buttons_frame.rowconfigure(1, weight=1)
        self.train_buttons_frame.rowconfigure(2, weight=1)
        self.train_buttons_frame.columnconfigure(0, weight=1)
        self.train_buttons_frame.columnconfigure(1, weight=1)



        self.train_btn = ctk.CTkButton(self.train_buttons_frame, text='Train Model',state="disabled", height=btn_height, width=80, command=self.train_model,
                                    font=("Inter", 14, "bold"))
        self.train_btn.grid(row=0, column=0, columnspan=2, padx=5, pady=10, sticky='SEW')
        self.load_model = ctk.CTkButton(self.train_buttons_frame, text='Load Model', height=btn_height, width=80,
                                    font=("Inter", 14, "bold"))
        self.load_model.grid(row=1, column=1, columnspan=1, padx=5, pady=10, sticky='SEW')
        self.save_model = ctk.CTkButton(self.train_buttons_frame,state="disabled", text='Save Model', height=btn_height, width=80,
                                                                            font=("Inter", 14, "bold"))
        self.save_model.grid(row=1, column=0, columnspan=1, padx=5, pady=10, sticky='SEW')

        #this will need to be made into a formatted string var so that it can show the training accuracy etc, alterntaively use multiple labels here (probably easier)
        self.training_label = ctk.CTkLabel(master=self.train_buttons_frame, text="Training Status: \nModel has not yet been trained", justify = "left", font=("Inter", 10, "italic"))
        self.training_label.grid(row=2, column=0, columnspan=2, padx=5, pady=10, sticky='SW')


        #add widgets to right frame
        #excel viewer widget
        self.right_buttons_frame = ctk.CTkFrame(master=self.train_frame_right, corner_radius=10, fg_color=LABEL_COLOR)
        self.right_buttons_frame.grid(row=0, column=0, padx=5, pady=5, sticky='NSEW')
        self.right_buttons_frame.rowconfigure(0, weight=1)
        self.right_buttons_frame.columnconfigure(0, weight=1)
        self.right_buttons_frame.columnconfigure(1, weight=1)
        self.right_buttons_frame.columnconfigure(2, weight=1)
        self.right_buttons_frame.columnconfigure(3, weight=1)

        self.filter_btn = (
            ctk.CTkButton(self.right_buttons_frame, text='Filter', height=btn_height,state="disabled", width=80, command=self.filter_data,
                                    font=("Inter", 14, "bold")))
        self.filter_btn.grid(row=0, column=0, columnspan=1, padx=5, pady=10, sticky='NEW')
        self.zeros_btn = (
            ctk.CTkButton(self.right_buttons_frame, text='Remove Zeros',state="disabled", height=btn_height,command=self.remove_zeros, width=80,
                                    font=("Inter", 14, "bold")))
        self.zeros_btn.grid(row=0, column=1, columnspan=1, padx=5, pady=10, sticky='NEW')
        self.reset_btn = (ctk.CTkButton(self.right_buttons_frame, text='Reset',state="disabled", height=btn_height, width=80, command=self.reset_data,
                                    font=("Inter", 14, "bold")))
        self.reset_btn.grid(row=0, column=2, columnspan=1, padx=5, pady=10, sticky='NEW')

        self.store_btn = (ctk.CTkButton(self.right_buttons_frame, text='Store Data',state="disabled", height=btn_height, width=80, command=self.store_data,
                                    font=("Inter", 14, "bold")))
        self.store_btn.grid(row=0, column=3, columnspan=1, padx=5, pady=10, sticky='NEW')

        #add a canvas widget to row 1 of the right frame
        self.canvas_frame = ctk.CTkFrame(master=self.train_frame_right, corner_radius=10, fg_color='red', width=600, height=200)
        self.canvas_frame.grid(row=2, column=0, padx=5, pady=5, sticky='NEWS')
        self.canvas = Canvas(self.canvas_frame, bg=LABEL_COLOR, bd=0, highlightthickness=0, relief='ridge', width=600, height=200)
        #self.canvas = Canvas(self.canvas_frame, bg='red', bd=0, highlightthickness=0, relief='ridge', width=600, height=200)
        #self.canvas.minsize(400, 200)
        #self.canvas.create_text(self.canvas.winfo_reqwidth() / 2, self.canvas.winfo_reqheight() / 2, fill="white",
        #                        font=("Inter", 12, "bold"),
        #                        anchor="center")
        self.canvas.pack(expand=True, fill="both")
        #place dnd image on the canvas
        self.image = Image.open("images/Gradebook_image1.png")
        self.original_image=self.image
        self.after_idle(self.update_canvas)



        self.canvas.drop_target_register(DND_ALL)
        self.canvas.dnd_bind("<<Drop>>", self.get_path)
        self.canvas.bind("<Configure>", self.on_resize)



        self.data_tree=ttk.Treeview(master = self.train_frame_right)

        #self.scaleFactor = ctypes.windll.shcore.GetScaleFactorForDevice(0) / 100
        self.scale_factor=0
        self.data_loaded = False
        self.get_scale_factor()
        #self.scaleFactor=1
        print(self.scale_factor)
        self.orig_scale=self.scale_factor

        self.checkboxes_frame_outer = None

        self.no_of_datafiles = 0
        self.stored_data_1 = None
        self.stored_data_2 = None
        self.stored_data_3 = None


    def file_dialog(self):
        #get path after an import button has been clicked ( no curly brackets with this method)
        self.file_path = filedialog.askopenfilename(filetypes=[("xlsx files", "*.xlsx")])
        # extract the filename from the full file path so it can be shown in the GUI
        self.filename = self.file_path.split("/")[-1]
        self.filename_var.set(self.filename)
        self.file_open()

    def reset_data(self):
        # get path after an import button has been clicked ( no curly brackets with this method)
        self.df = self.orig_df
        self.load_data()
        print("data reset")
    def store_data(self):
        # get path after an import button has been clicked ( no curly brackets with this method)
        self.no_of_datafiles = self.no_of_datafiles + 1
        self.string_no_of_datafiles = str(self.no_of_datafiles)
        self.no_of_datafiles_SV.set(self.string_no_of_datafiles)

        if self.no_of_datafiles == 1:
            self.stored_data_1 = self.df
            self.check_store_1.set(True)
            self.file_1_name.set(self.filename)
            self.file_1_string = self.file_1_name.get()
            #self.check_file_1.grid(row=1, column=0, padx=5, pady=5, sticky='NW')
            CTkToolTip(self.check_file_1,
                       message=self.file_1_string,
                       delay=0.5, alpha=1, wraplength=200)
            self.train_btn.configure(state="normal")
            self.remove_data_tree()
            self.canvas_frame.grid(row=2, column=0, padx=5, pady=5, sticky='NEWS')
            self.canvas.grid(row=0, column=0, padx=5, pady=5, sticky='NEWS')
            self.canvas.delete(self.image_on_canvas)
            self.canvas.pack(expand=True, fill="both")
            # place dnd image on the canvas
            self.image = Image.open("images/Gradebook_image2.png")
            self.update_canvas()

        elif self.no_of_datafiles == 2:
            self.stored_data_2 = self.df
            self.check_store_2.set(True)
            self.file_2_name.set(self.filename)
            self.check_file_2.grid(row=2, column=0,columnspan=2, padx=5, pady=5, sticky='NW')
            self.file_2_string = self.file_2_name.get()
            CTkToolTip(self.check_file_2,
                       message=self.file_2_string,
                       delay=0.5, alpha=1, wraplength=200)
            self.remove_data_tree()
            self.canvas_frame.grid(row=2, column=0, padx=5, pady=5, sticky='NEWS')
            self.canvas.grid(row=0, column=0, padx=5, pady=5, sticky='NEWS')
            self.canvas.delete(self.image_on_canvas)
            self.canvas.pack(expand=True, fill="both")
            # place dnd image on the canvas
            self.image = Image.open("images/Gradebook_image2.png")
            #self.after_idle(self.update_canvas)
            self.update_canvas()
        elif self.no_of_datafiles == 3:
            self.stored_data_3 = self.df
            self.check_store_3.set(True)
            self.file_3_name.set(self.filename)
            self.check_file_3.grid(row=3, column=0,columnspan=2, padx=5, pady=5, sticky='NW')
            self.file_3_string = self.file_3_name.get()
            CTkToolTip(self.check_file_3,
                       message=self.file_3_string,
                       delay=0.5, alpha=1, wraplength=200)
            self.remove_data_tree()
            self.canvas_frame.grid(row=2, column=0, padx=5, pady=5, sticky='NEWS')
            self.canvas.grid(row=0, column=0, padx=5, pady=5, sticky='NEWS')
            self.canvas.delete(self.image_on_canvas)
            self.canvas.pack(expand=True, fill="both")
            # place dnd image on the canvas
            self.image = Image.open("images/Gradebook_image3.png")
            self.update_canvas()
        else:
            print("Max number of datafiles reached")
        print("data stored")

    def update_canvas(self):
        self.canvas_width = self.canvas.winfo_reqwidth()
        self.canvas_height = self.canvas.winfo_reqheight()
        self.original_image = self.image
        self.image = self.original_image.resize((self.canvas_width, self.canvas_height))
        self.photo = ImageTk.PhotoImage(image=self.image)
        self.image_on_canvas = self.canvas.create_image(int(self.canvas_width / 2), int(self.canvas_height / 2),
                                                        image=self.photo)

    def get_path(self, event):
        #Get the filepath from drag n drop image. remove the curly brackets
        self.file_path = event.data
        if self.file_path[0] == '{' and self.file_path[-1] == '}':
            self.file_path = self.file_path[1:-1]
        self.filename = self.file_path.split("/")[-1]
        self.filename_var.set(self.filename)
        self.file_open()


    def test_function(self):
        print("test")

    def train_model(self):
        self.save_model.configure(state="normal")



    def filter_data(self):
        # Create a copy of the column names
        columns = list(self.df.columns)

        # Iterate over the check_vars list and the column names together
        for check_var, column in zip(self.check_vars, columns):
            # If the checkbox is unchecked
            if not check_var.get():
                # Drop the column from the DataFrame
                self.df = self.df.drop(column, axis=1)

        # After filtering, you might want to reload the data
        self.load_data()
    def remove_zeros(self):

            # Calculate the threshold for number of zeros or '-' characters
        threshold = 0.7 * len(self.df.columns)

        # Define a function to apply to each row
        def check_row(row):
            # Count the number of zeros or '-' characters in the row
            zero_dash_count = sum((row == 0) | (row == '-'))
            # Return True if the count is less than the threshold (i.e., keep the row)
            # and False otherwise (i.e., remove the row)
            return zero_dash_count < threshold

        # Apply the function to each row and keep only the rows where the function returns True
        self.df = self.df[self.df.apply(check_row, axis=1)]
        print("zeros removed")
        self.load_data()


    def file_open(self):
        if self.filename:
            try:
                self.df=pd.read_excel(self.file_path)
                self.orig_df= self.df
                print("file opened")
                self.load_data()
            except ValueError:
                self.filename_var.set("Error", "The file you have chosen is invalid")
                return

    def load_data(self):
        #clear  old tree view
        self.clear_tree()
        self.canvas_frame.grid_forget()
        check_height=50
        # Create a new frame for the checkboxes
        #self.checkboxes_frame_outer = ctk.CTkFrame(self.train_frame_right, height =check_height, corner_radius=10, fg_color=MAIN_FRAME_COLOR,  border_width=0)
        self.checkboxes_frame_outer = ctk.CTkFrame(self.train_frame_right, height =check_height, corner_radius=10, fg_color=MAIN_FRAME_COLOR,  border_width=0)
        self.checkboxes_frame_outer.grid(row=1, column=0, columnspan =1, padx=5, pady=0, sticky='news')

        # Add a Canvas to this frame
        self.checkboxes_canvas = Canvas(self.checkboxes_frame_outer, height =check_height, bg=LABEL_COLOR, borderwidth=0, bd=0, relief='ridge', highlightthickness=0)
        self.checkboxes_canvas.pack(side='left', fill='both', expand=True)

        # Create a new frame for the checkboxes and add it to the Canvas
        self.checkboxes_frame = ctk.CTkFrame(self.checkboxes_canvas, height =check_height, corner_radius=10, fg_color=LABEL_COLOR, border_width=0)

        #adjusting 0,0 here might make it render quicker
        self.checkboxes_canvas.create_window((0, 0), window=self.checkboxes_frame, anchor='sw')

        #set up new tree view
        self.data_tree["column"] = list(self.df.columns)
        #replace below with empty string to remove headings line
        self.data_tree["show"] = ""
        #loop through column lists and add them to the tree view
        self.check_vars = []
        i=0
        col_width=120
        scaled_width=int(col_width*self.scale_factor)
        for column in self.data_tree["columns"]:

            check_var = BooleanVar()
            num_chars=15
            short_col= column[:num_chars]
            self.check_vars.append(check_var)

            self.check_ind_frame = ctk.CTkFrame(self.checkboxes_frame, height=check_height, width=col_width)
            self.check_ind_frame.grid(row=0, column=i, padx=(0,0), pady=0,ipadx=0, ipady=0,  sticky="sew")
            self.check_ind_frame.pack_propagate(False)
            self.check_ind_frame.grid_propagate(False)
            #font = ctkFont.Font(font=widget.cget("font"))
            #char_width= font.measure("0");
            check_button = ctk.CTkCheckBox(self.check_ind_frame, variable=check_var, command=self.on_check, text=short_col, width=15, height=15, font=("Inter", 10))
            #font = tkFont.Font(font=check_button.cget("font"))
            #char_width = font.measure("0")
            #check_button_width_pixels = char_width * num_chars

            #code below to enable tooltips
            CTkToolTip(check_button,
                       message=column,
                       delay=0.5, alpha=1, wraplength=200)

            #check_button.grid(row=0, column=0, sticky="EW")
            #check_button.grid(row=0, column=i,padx=(0,0), pady=0, sticky="")
            #check_button.pack(side='bottom', padx=(5,0), expand=True, fill='both')
            check_button.pack(side='bottom', padx=(5,0), expand=True, fill='both')
            check_button.pack_propagate(False)
            self.data_tree.column(column, stretch=False, width=scaled_width)
            self.data_tree.heading(column, command=lambda: "break", text=column)
            self.data_tree.bind('<Button-1>', self.do_nothing)
            # get width of the current column
            #col_width = self.data_tree.column(column)['width']
            #check_button.pack(side='top')
            i=i+1

        #put the data in the treeview
        df_rows = self.df.to_numpy().tolist()
        for row in df_rows:
            #check_var = BooleanVar()
            #self.check_vars.append(check_var)
            #check_button = ctk.CTkCheckBox(self.data_tree, variable=check_var, command=self.on_check)
            #self.data_tree.insert("", "end", values=(check_button,) + tuple(row))
            self.data_tree.insert("", "end", values=row)

        #self.data_tree.configure(show="tree")
        #self.data_tree.configure(height=100)
        self.data_tree.grid(row=2, column=0, columnspan=1, padx=5, pady=5, ipadx=0, ipady=0, sticky='NEWS')
        self.data_tree.grid_propagate(False)
        # Add these lines where you create your Treeview widget
        self.yscrollbar = ctk.CTkScrollbar(self.train_frame_right, command=self.data_tree.yview)
        self.data_tree.configure(yscrollcommand=self.yscrollbar.set)

        # Place the scrollbar next to the Treeview
        #self.data_tree.grid(row=1, column=0, sticky='nsew')
        self.yscrollbar.grid(row=2, column=1, sticky='ns')

        self.x_scrollbar = ctk.CTkScrollbar(self.train_frame_right, command=self.multi_scroll, orientation="horizontal")

        self.x_scrollbar.grid(row=3, column=0, sticky='ew')
        self.data_tree.configure(xscrollcommand=self.x_scrollbar.set)
        self.checkboxes_canvas.configure(xscrollcommand=self.x_scrollbar.set)

        # Update the scroll region of the Canvas
        self.checkboxes_frame.update_idletasks()
        self.checkboxes_canvas.configure(scrollregion=self.checkboxes_canvas.bbox('all'))

        self.after_idle(lambda: self.x_scrollbar.set(0, 0))
        self.after_idle(lambda: self.data_tree.xview_moveto(0))
        self.after_idle(lambda: self.checkboxes_canvas.xview_moveto(0))

        self.data_loaded = True

        self.filter_btn.configure(state="normal")
        self.zeros_btn.configure(state="normal")
        self.reset_btn.configure(state="normal")
        self.store_btn.configure(state="normal")

        # Update the scroll region of the Canvas
        #self.checkboxes_frame.update_idletasks()
        #self.checkboxes_canvas.configure(scrollregion=self.checkboxes_canvas.bbox('all'))



    def on_resize(self, event):
        # Get the current scale factor
        #print("resized")
        self.get_scale_factor()
        if self.scale_factor != self.orig_scale:
            print("Scale Factor Changed")
            self.orig_scale = self.scale_factor

            #The following code will work to reopen and rescale the file for a new monitor DPI, but it is quite slow moving between monitors
            #current issue that the checkboxes stop scrolling fully
            if self.data_loaded==True:
                self.remove_data_tree()
                self.file_open()

        self.after_idle(self.update_image)

    def remove_data_tree(self):
        self.checkboxes_frame_outer.grid_forget()
        self.data_tree.grid_forget()
        self.x_scrollbar.grid_forget()
        self.yscrollbar.grid_forget()
    def update_image(self):
        self.canvas_width = self.canvas_frame.winfo_width()
        self.canvas_height = self.canvas_frame.winfo_height()

        self.image = self.original_image.resize((self.canvas_width, self.canvas_height))
        self.photo = ImageTk.PhotoImage(image=self.image)
        self.image_on_canvas = self.canvas.create_image(int(self.canvas_width / 2), int(self.canvas_height / 2),
                                                        image=self.photo)

    def get_scale_factor(self):
        winID = ctypes.windll.user32.GetForegroundWindow()

        MONITOR_DEFAULTTONULL = 0
        MONITOR_DEFAULTTOPRIMARY = 1
        MONITOR_DEFAULTTONEAREST = 2
        self.MDT_EFFECTIVE_DPI=0

        monitorID = ctypes.windll.user32.MonitorFromWindow(winID, MONITOR_DEFAULTTONEAREST)
        dpiX = ctypes.c_uint()
        dpiY = ctypes.c_uint()
        ctypes.windll.shcore.GetDpiForMonitor(monitorID, self.MDT_EFFECTIVE_DPI, ctypes.byref(dpiX), ctypes.byref(dpiY))

        self.scale_factor  = dpiX.value / 96
        #print("Current Scale Factor: ", self.scale_factor)

    def clear_tree(self):
        self.data_tree.delete(*self.data_tree.get_children())

    def multi_scroll(self, *args):

        self.data_tree.xview(*args)
        self.checkboxes_canvas.xview(*args)
        print("scrolling")



    def do_nothing(event, self):
        return "break"
    def on_check(self):
        # This function will be called when a checkbox is clicked
        # You can add your own logic here
        for i, check_var in enumerate(self.check_vars):
            print(f"Checkbox {i} is {'checked' if check_var.get() else 'not checked'}")












if __name__ == '__main__':
    app = MainApp()
    app.mainloop()