# using version 5.22 of customtkinter
import customtkinter as ctk
from settings import *
from tkinterdnd2 import TkinterDnD, DND_ALL
from tkinter import filedialog, Canvas, StringVar, ttk, BooleanVar, DISABLED
import pandas as pd
from tkinter import font as tkFont
import openpyxl
from CTkToolTip import *
from PIL import Image, ImageTk, ImageOps

# importing pyautogui will cause the app to not rescale when moving between windows, poor resolution, but stable columns!
import pyautogui
# from screeninfo import get_monitors
import ctypes

# load color theme
ctk.set_default_color_theme('themes\phystek_colours.json')


# ctk.set_default_color_theme('dark-blue')


class MainApp(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self):

        # setup
        super().__init__()
        ctk.set_appearance_mode("dark")
        w, h = 800, 450  # Width and height.
        x, y = 200, 200  # Screen position.
        self.geometry('%dx%d+%d+%d' % (w, h, x, y))
        self.minsize(800, 450)
        self.title('AI Grade Prediction')
        self.TkdndVersion = TkinterDnD._require(self)

        self.resizable(True, True)  # Disable resizing of the window
        # trigger resize function on window configuration change

        self.bind("<Configure>", self.on_resize)

        # canvas data (not currently used)
        self.canvas_width = 0
        self.canvas_height = 0
        btn_height = 26
        btn_width = 135
        self.check_vars = []

        # main frame here, holds all of the other frames with their widgets
        frame_main = ctk.CTkFrame(master=self, corner_radius=10, fg_color=MAIN_FRAME_COLOR, bg_color=MAIN_FRAME_COLOR)
        frame_main.pack(padx=0, pady=0, fill="both", expand=True)

        # add prediction and training tabs to the main frame
        self.main_tabs = ctk.CTkTabview(master=frame_main, corner_radius=10, fg_color=LABEL_COLOR,
                                        bg_color=MAIN_FRAME_COLOR)
        self.main_tabs._segmented_button.grid(sticky="W")

        self.main_tabs.pack(padx=5, pady=(0, 5), ipadx=0, ipady=0, fill="both", expand=True, side="left")
        self.main_tabs.add("Train Model")
        self.main_tabs.add("Predict Grades")
        self.main_tabs._segmented_button.grid(sticky="W")

        # add frame to the training tab
        self.train_frame = ctk.CTkFrame(master=self.main_tabs.tab("Train Model"), corner_radius=10,
                                        fg_color=MAIN_FRAME_COLOR)
        self.train_frame.pack(padx=0, pady=0, fill="both", expand=True)
        self.train_frame.rowconfigure(0, weight=1)
        self.train_frame.columnconfigure(0, weight=1)
        self.train_frame.columnconfigure(1, weight=4)

        # add frame to the prediction tab
        self.predict_frame = ctk.CTkFrame(master=self.main_tabs.tab("Predict Grades"), corner_radius=10,
                                          fg_color=MAIN_FRAME_COLOR)
        self.predict_frame.pack(padx=0, pady=0, fill="both", expand=True)
        self.predict_frame.rowconfigure(0, weight=1)
        self.predict_frame.columnconfigure(0, weight=1)
        self.predict_frame.columnconfigure(1, weight=4)

        # make left and right frames for the train gui
        self.train_frame_left = ctk.CTkFrame(master=self.train_frame, corner_radius=10, fg_color=LABEL_COLOR, width=200)
        self.train_frame_left.grid(row=0, column=0, padx=5, pady=5, sticky='NSEW')
        self.train_frame_left.rowconfigure(0, weight=1)
        self.train_frame_left.rowconfigure(1, weight=1)
        self.train_frame_left.rowconfigure(2, weight=10)
        self.train_frame_left.columnconfigure(0, weight=1)
        self.train_frame_left.grid_propagate(False)

        self.train_frame_right = ctk.CTkFrame(master=self.train_frame, corner_radius=10, fg_color=LABEL_COLOR,
                                              width=400)
        self.train_frame_right.grid(row=0, column=1, padx=5, pady=5, sticky='NSEW')
        self.train_frame_right.rowconfigure(0, weight=1)
        self.train_frame_right.rowconfigure(1, weight=20)
        self.train_frame_right.columnconfigure(0, weight=1)
        self.train_frame_right.grid_propagate(False)

        # make left and right frames for the predict gui
        self.predict_frame_left = ctk.CTkFrame(master=self.predict_frame, corner_radius=10, fg_color=LABEL_COLOR,
                                               width=200)
        self.predict_frame_left.grid(row=0, column=0, padx=5, pady=5, sticky='NSEW')
        self.predict_frame_left.rowconfigure(0, weight=1)
        self.predict_frame_left.rowconfigure(1, weight=1)
        self.predict_frame_left.rowconfigure(2, weight=10)
        self.predict_frame_left.columnconfigure(0, weight=1)
        self.predict_frame_left.grid_propagate(False)

        self.predict_frame_right = ctk.CTkFrame(master=self.predict_frame, corner_radius=10, fg_color=LABEL_COLOR,
                                                width=400)
        self.predict_frame_right.grid(row=0, column=1, padx=5, pady=5, sticky='NSEW')
        self.predict_frame_right.rowconfigure(0, weight=1)
        self.predict_frame_right.rowconfigure(1, weight=20)
        self.predict_frame_right.columnconfigure(0, weight=1)
        self.predict_frame_right.grid_propagate(False)

        # import widgets frame for training
        self.import_frame = ctk.CTkFrame(master=self.train_frame_left, corner_radius=10, fg_color=LABEL_COLOR)
        self.import_frame.grid(row=0, column=0, padx=5, pady=5, sticky='NEW')
        self.import_frame.rowconfigure(0, weight=1)
        self.import_frame.rowconfigure(1, weight=1)
        self.import_frame.columnconfigure(0, weight=1)

        # import widgets frame for prediction
        self.import_frame_p = ctk.CTkFrame(master=self.predict_frame_left, corner_radius=10, fg_color=LABEL_COLOR)
        self.import_frame_p.grid(row=0, column=0, padx=5, pady=5, sticky='NEW')
        self.import_frame_p.rowconfigure(0, weight=1)
        self.import_frame_p.rowconfigure(1, weight=1)
        self.import_frame_p.columnconfigure(0, weight=1)

        # add widgets to left predict frame
        self.import_file_btn_p = (
            ctk.CTkButton(self.import_frame_p, text='Import Gradebook File', height=btn_height,
                          command=self.file_dialog,
                          font=("Inter", 14, "bold")))
        self.import_file_btn_p.grid(row=0, column=0, columnspan=1, padx=5, pady=10, sticky='NEW')
        self.filename_var_p = StringVar(value="Filename: No files loaded")
        self.entryWidget_p = ctk.CTkEntry(master=self.import_frame_p, font=("Inter", 10, "italic"),
                                          textvariable=self.filename_var_p, height=22)
        self.entryWidget_p.grid(row=1, column=0, columnspan=1, padx=5, pady=3, ipadx=0, ipady=0, sticky='NEW')

        self.entryWidget_p.drop_target_register(DND_ALL)
        self.entryWidget_p.dnd_bind("<<Drop>>", self.get_path_p)

        self.stored_data_frame_p = ctk.CTkFrame(master=self.predict_frame_left, corner_radius=10, fg_color=LABEL_COLOR)
        self.stored_data_frame_p.grid(row=1, column=0, padx=5, pady=5, sticky='NEW')
        self.stored_data_frame_p.rowconfigure(0, weight=1)
        self.stored_data_frame_p.rowconfigure(1, weight=1)
        self.stored_data_frame_p.columnconfigure(0, weight=1)
        self.stored_data_frame_p.columnconfigure(1, weight=1)

        self.stored_files_label_p = ctk.CTkLabel(master=self.stored_data_frame_p, text="Number of files stored:",
                                                 font=("Inter", 12))
        self.stored_files_label_p.grid(row=0, column=0, padx=5, pady=5, sticky='NW')
        self.no_of_datafiles_SV_p = StringVar(value="0")
        self.num_files_label_p = ctk.CTkLabel(master=self.stored_data_frame_p, textvariable=self.no_of_datafiles_SV_p,
                                              font=("Inter", 12, "bold"))
        self.num_files_label_p.grid(row=0, column=1, padx=1, pady=5, sticky='NW')

        self.check_store_1_p = BooleanVar(value=False)
        self.file_1_name_p = StringVar(value="No Files Stored")
        self.check_file_1_p = ctk.CTkCheckBox(master=self.stored_data_frame_p, variable=self.check_store_1_p,
                                              state=DISABLED,
                                              textvariable=self.file_1_name_p, width=12, height=12, font=("Inter", 10))
        self.check_file_1_p.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky='NW')

        self.predict_buttons_frame = ctk.CTkFrame(master=self.predict_frame_left, corner_radius=10,
                                                  fg_color=LABEL_COLOR)
        self.predict_buttons_frame.grid(row=2, column=0, padx=5, pady=5, sticky='SEW')
        self.predict_buttons_frame.rowconfigure(0, weight=1)
        self.predict_buttons_frame.rowconfigure(1, weight=1)
        self.predict_buttons_frame.rowconfigure(2, weight=1)
        self.predict_buttons_frame.columnconfigure(0, weight=1)
        self.predict_buttons_frame.columnconfigure(1, weight=1)

        self.predict_btn = ctk.CTkButton(self.predict_buttons_frame, text='Predict Grades', state="disabled",
                                         height=btn_height, width=80, command=self.train_model,
                                         font=("Inter", 14, "bold"))
        self.predict_btn.grid(row=0, column=0, columnspan=2, padx=5, pady=10, sticky='SEW')
        self.save_predict = ctk.CTkButton(self.predict_buttons_frame, text='Save Prediction', height=btn_height,
                                          width=80,
                                          font=("Inter", 14, "bold"))

        self.save_predict.grid(row=1, column=0, columnspan=2, padx=5, pady=10, sticky='SEW')

        # this will need to be made into a formatted string var so that it can show the training accuracy etc, alterntaively use multiple labels here (probably easier)
        self.predict_label = ctk.CTkLabel(master=self.predict_buttons_frame,
                                          text="Prediction Status: \nPrediction has not yet been generated",
                                          justify="left",
                                          font=("Inter", 10, "italic"))
        self.predict_label.grid(row=2, column=0, columnspan=2, padx=5, pady=10, sticky='SW')

        # add widgets to left training frame
        self.import_file_btn = (
            ctk.CTkButton(self.import_frame, text='Import Gradebook File', height=btn_height, command=self.file_dialog,
                          font=("Inter", 14, "bold")))
        self.import_file_btn.grid(row=0, column=0, columnspan=1, padx=5, pady=10, sticky='NEW')
        self.filename_var = StringVar(value="Filename: No files loaded")
        self.entryWidget = ctk.CTkEntry(master=self.import_frame, font=("Inter", 10, "italic"),
                                        textvariable=self.filename_var, height=22)
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

        self.stored_files_label = ctk.CTkLabel(master=self.stored_data_frame, text="Number of files stored:",
                                               font=("Inter", 12))
        self.stored_files_label.grid(row=0, column=0, padx=5, pady=5, sticky='NW')
        self.no_of_datafiles_SV = StringVar(value="0")
        self.num_files_label = ctk.CTkLabel(master=self.stored_data_frame, textvariable=self.no_of_datafiles_SV,
                                            font=("Inter", 12, "bold"))
        self.num_files_label.grid(row=0, column=1, padx=1, pady=5, sticky='NW')

        self.check_store_1 = BooleanVar(value=False)
        self.file_1_name = StringVar(value="No Files Stored")
        self.check_file_1 = ctk.CTkCheckBox(master=self.stored_data_frame, variable=self.check_store_1, state=DISABLED,
                                            textvariable=self.file_1_name, width=12, height=12, font=("Inter", 10))
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

        self.train_btn = ctk.CTkButton(self.train_buttons_frame, text='Train Model', state="disabled",
                                       height=btn_height, width=80, command=self.train_model,
                                       font=("Inter", 14, "bold"))
        self.train_btn.grid(row=0, column=0, columnspan=2, padx=5, pady=10, sticky='SEW')
        self.load_model = ctk.CTkButton(self.train_buttons_frame, text='Load Model', height=btn_height, width=80,
                                        font=("Inter", 14, "bold"))
        self.load_model.grid(row=1, column=1, columnspan=1, padx=5, pady=10, sticky='SEW')
        self.save_model = ctk.CTkButton(self.train_buttons_frame, state="disabled", text='Save Model',
                                        height=btn_height, width=80,
                                        font=("Inter", 14, "bold"))
        self.save_model.grid(row=1, column=0, columnspan=1, padx=5, pady=10, sticky='SEW')

        # this will need to be made into a formatted string var so that it can show the training accuracy etc, alterntaively use multiple labels here (probably easier)
        self.training_label = ctk.CTkLabel(master=self.train_buttons_frame,
                                           text="Training Status: \nModel has not yet been trained", justify="left",
                                           font=("Inter", 10, "italic"))
        self.training_label.grid(row=2, column=0, columnspan=2, padx=5, pady=10, sticky='SW')

        # add widgets to right frame
        # excel viewer widget
        self.right_buttons_frame = ctk.CTkFrame(master=self.train_frame_right, height=60, width=300, corner_radius=10,
                                                fg_color=LABEL_COLOR)
        self.right_buttons_frame.grid(row=0, column=0, padx=5, pady=5, sticky='EW')
        self.right_buttons_frame.rowconfigure(0, weight=1)
        self.right_buttons_frame.columnconfigure(0, weight=1)
        self.right_buttons_frame.columnconfigure(1, weight=1)
        self.right_buttons_frame.columnconfigure(2, weight=1)
        self.right_buttons_frame.columnconfigure(3, weight=1)
        # self.right_buttons_frame.grid_propagate(False)

        self.filter_btn = (
            ctk.CTkButton(self.right_buttons_frame, text='Filter', height=btn_height, state="disabled", width=80,
                          command=self.filter_data,
                          font=("Inter", 14, "bold")))
        self.filter_btn.grid(row=0, column=0, columnspan=1, padx=10, pady=10, sticky='NEW')
        self.zeros_btn = (
            ctk.CTkButton(self.right_buttons_frame, text='Remove Zeros', state="disabled", height=btn_height,
                          command=self.remove_zeros, width=80,
                          font=("Inter", 14, "bold")))
        self.zeros_btn.grid(row=0, column=1, columnspan=1, padx=10, pady=10, sticky='NEW')
        self.reset_btn = (
            ctk.CTkButton(self.right_buttons_frame, text='Reset', state="disabled", height=btn_height, width=80,
                          command=self.reset_data,
                          font=("Inter", 14, "bold")))
        self.reset_btn.grid(row=0, column=2, columnspan=1, padx=10, pady=10, sticky='NEW')

        self.store_btn = (
            ctk.CTkButton(self.right_buttons_frame, text='Store Data', state="disabled", height=btn_height, width=80,
                          command=self.store_data,
                          font=("Inter", 14, "bold")))
        self.store_btn.grid(row=0, column=3, columnspan=1, padx=10, pady=10, sticky='NEW')

        self.dynamic_right_frame = ctk.CTkFrame(master=self.train_frame_right, corner_radius=10, fg_color=LABEL_COLOR)
        self.dynamic_right_frame.grid(row=1, column=0, padx=5, pady=5, sticky='NEWS')
        self.dynamic_right_frame.rowconfigure(0, weight=1)
        self.dynamic_right_frame.rowconfigure(1, weight=20)
        self.dynamic_right_frame.rowconfigure(2, weight=1)
        self.dynamic_right_frame.columnconfigure(0, weight=1)
        # self.dynamic_right_frame.grid_propagate(False)

        # add a canvas widget to row 1 of the right frame
        self.canvas_frame = ctk.CTkFrame(master=self.train_frame_right, corner_radius=10, fg_color='red', width=600,
                                         height=200)
        self.canvas_frame.grid(row=1, column=0, padx=5, pady=5, sticky='NEWS')
        self.canvas = Canvas(self.canvas_frame, bg=LABEL_COLOR, bd=0, highlightthickness=0, relief='ridge', width=600,
                             height=200)
        self.canvas.pack(expand=True, fill="both")
        # place dnd image on the canvas
        self.image = Image.open("images/Gradebook_image1.png")
        self.original_image = self.image
        self.after_idle(self.update_canvas)

        # bind drang and drop to the canvas
        self.canvas.drop_target_register(DND_ALL)
        self.canvas.dnd_bind("<<Drop>>", self.get_path)
        self.canvas.bind("<Configure>", self.on_resize)

        # set up data tree for incoming file data
        self.data_tree = ttk.Treeview(master=self.dynamic_right_frame)

        # add widgets to right prediction frame
        # excel viewer widget
        self.right_buttons_frame_p = ctk.CTkFrame(master=self.predict_frame_right, height=60, width=300,
                                                  corner_radius=10,
                                                  fg_color=LABEL_COLOR)
        self.right_buttons_frame_p.grid(row=0, column=0, padx=5, pady=5, sticky='EW')
        self.right_buttons_frame_p.rowconfigure(0, weight=1)
        self.right_buttons_frame_p.columnconfigure(0, weight=1)
        self.right_buttons_frame_p.columnconfigure(1, weight=1)
        self.right_buttons_frame_p.columnconfigure(2, weight=1)
        self.right_buttons_frame_p.columnconfigure(3, weight=1)
        # self.right_buttons_frame.grid_propagate(False)

        self.filter_btn_p = (
            ctk.CTkButton(self.right_buttons_frame_p, text='Filter', height=btn_height, state="disabled", width=80,
                          command=self.filter_data_p,
                          font=("Inter", 14, "bold")))
        self.filter_btn_p.grid(row=0, column=0, columnspan=1, padx=10, pady=10, sticky='NEW')
        self.zeros_btn_p = (
            ctk.CTkButton(self.right_buttons_frame_p, text='Remove Zeros', state="disabled", height=btn_height,
                          command=self.remove_zeros_p, width=80,
                          font=("Inter", 14, "bold")))
        self.zeros_btn_p.grid(row=0, column=1, columnspan=1, padx=10, pady=10, sticky='NEW')
        self.reset_btn_p = (
            ctk.CTkButton(self.right_buttons_frame_p, text='Reset', state="disabled", height=btn_height, width=80,
                          command=self.reset_data_p,
                          font=("Inter", 14, "bold")))
        self.reset_btn_p.grid(row=0, column=2, columnspan=1, padx=10, pady=10, sticky='NEW')

        self.store_btn_p = (
            ctk.CTkButton(self.right_buttons_frame_p, text='Store Data', state="disabled", height=btn_height, width=80,
                          command=self.store_data_p,
                          font=("Inter", 14, "bold")))
        self.store_btn_p.grid(row=0, column=3, columnspan=1, padx=10, pady=10, sticky='NEW')

        self.dynamic_right_frame_p = ctk.CTkFrame(master=self.predict_frame_right, corner_radius=10,
                                                  fg_color=LABEL_COLOR)
        self.dynamic_right_frame_p.grid(row=1, column=0, padx=5, pady=5, sticky='NEWS')
        self.dynamic_right_frame_p.rowconfigure(0, weight=1)
        self.dynamic_right_frame_p.rowconfigure(1, weight=20)
        self.dynamic_right_frame_p.rowconfigure(2, weight=1)
        self.dynamic_right_frame_p.columnconfigure(0, weight=1)
        # self.dynamic_right_frame.grid_propagate(False)

        # add a canvas widget to row 1 of the right frame
        self.canvas_frame_p = ctk.CTkFrame(master=self.predict_frame_right, corner_radius=10, fg_color='red', width=600,
                                           height=200)
        self.canvas_frame_p.grid(row=1, column=0, padx=5, pady=5, sticky='NEWS')
        self.canvas_p = Canvas(self.canvas_frame_p, bg=LABEL_COLOR, bd=0, highlightthickness=0, relief='ridge',
                               width=600,
                               height=200)
        self.canvas_p.pack(expand=True, fill="both")
        # place dnd image on the canvas
        self.image_p = Image.open("images/Gradebook_image_current_lo.png")
        self.original_image_p = self.image_p
        self.after_idle(self.update_canvas)

        # bind drang and drop to the canvas
        self.canvas_p.drop_target_register(DND_ALL)
        self.canvas_p.dnd_bind("<<Drop>>", self.get_path_p)
        self.canvas_p.bind("<Configure>", self.on_resize)

        # set up data tree for incoming file data
        self.data_tree_p = ttk.Treeview(master=self.dynamic_right_frame_p)

        # set up variables
        self.scale_factor = 0
        self.data_loaded = False
        self.get_scale_factor()
        print(self.scale_factor)
        self.orig_scale = self.scale_factor
        self.checkboxes_frame_outer = None
        self.no_of_datafiles = 0
        self.stored_data_1 = None
        self.stored_data_2 = None
        self.stored_data_3 = None
        self.no_of_datafiles_p = 0
        self.stored_data_1_p = None

    def file_dialog(self):
        # get path after an import button has been clicked ( no curly brackets with this method)
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

    def reset_data_p(self):
        # get path after an import button has been clicked ( no curly brackets with this method)
        self.df_p = self.orig_df_p
        self.load_data()
        print("data reset")

    def store_data(self):
        # get path after an import button has been clicked ( no curly brackets with this method)
        self.no_of_datafiles = self.no_of_datafiles + 1
        self.string_no_of_datafiles = str(self.no_of_datafiles)
        self.no_of_datafiles_SV.set(self.string_no_of_datafiles)
        self.disable_data_buttons()

        if self.no_of_datafiles == 1:
            self.stored_data_1 = self.df
            self.check_store_1.set(True)
            self.file_1_name.set(self.filename)
            self.file_1_string = self.file_1_name.get()
            # self.check_file_1.grid(row=1, column=0, padx=5, pady=5, sticky='NW')
            CTkToolTip(self.check_file_1,
                       message=self.file_1_string,
                       delay=0.5, alpha=1, wraplength=200)
            self.train_btn.configure(state="normal")
            self.remove_data_tree()
            self.canvas_frame.grid(row=1, column=0, padx=5, pady=5, sticky='NEWS')
            # self.canvas.grid(row=0, column=0, padx=5, pady=5, sticky='NEWS')
            self.canvas.delete(self.image_on_canvas)

            # place dnd image on the canvas
            self.image = Image.open("images/Gradebook_image2.png")
            self.original_image = self.image
            self.update_canvas()
            self.canvas.pack(expand=True, fill="both")

        elif self.no_of_datafiles == 2:
            self.stored_data_2 = self.df
            self.check_store_2.set(True)
            self.file_2_name.set(self.filename)
            self.check_file_2.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky='NW')
            self.file_2_string = self.file_2_name.get()
            CTkToolTip(self.check_file_2,
                       message=self.file_2_string,
                       delay=0.5, alpha=1, wraplength=200)
            self.remove_data_tree()
            self.canvas_frame.grid(row=1, column=0, padx=5, pady=5, sticky='NEWS')
            # self.canvas.grid(row=0, column=0, padx=5, pady=5, sticky='NEWS')
            self.canvas.delete(self.image_on_canvas)

            # place dnd image on the canvas
            self.image = Image.open("images/Gradebook_image2.png")
            self.original_image = self.image
            # self.after_idle(self.update_canvas)
            self.update_canvas()
            self.canvas.pack(expand=True, fill="both")
        elif self.no_of_datafiles == 3:
            self.stored_data_3 = self.df
            self.check_store_3.set(True)
            self.file_3_name.set(self.filename)
            self.check_file_3.grid(row=3, column=0, columnspan=2, padx=5, pady=5, sticky='NW')
            self.file_3_string = self.file_3_name.get()
            CTkToolTip(self.check_file_3,
                       message=self.file_3_string,
                       delay=0.5, alpha=1, wraplength=200)
            self.remove_data_tree()
            self.canvas_frame.grid(row=1, column=0, padx=5, pady=5, sticky='NEWS')
            # self.canvas.grid(row=0, column=0, padx=5, pady=5, sticky='NEWS')

            # place dnd image on the canvas
            self.image = Image.open("images/Gradebook_image3.png")
            self.original_image = self.image
            self.canvas.delete(self.image_on_canvas)
            self.update_canvas()
            self.update_idletasks()
            self.canvas.pack(expand=True, fill="both")
        else:
            print("Max number of datafiles reached")
        print("data stored")

    def store_data_p(self):
        # get path after an import button has been clicked ( no curly brackets with this method)
        self.no_of_datafiles_p = self.no_of_datafiles_p + 1
        self.string_no_of_datafiles_p = str(self.no_of_datafiles_p)
        self.no_of_datafiles_SV_p.set(self.string_no_of_datafiles_p)
        self.disable_data_buttons_p()

        if self.no_of_datafiles == 1:
            self.stored_data_1_p = self.df_p
            self.check_store_1_p.set(True)
            self.file_1_name_p.set(self.filename_p)
            self.file_1_string_p = self.file_1_name_p.get()
            # self.check_file_1.grid(row=1, column=0, padx=5, pady=5, sticky='NW')
            CTkToolTip(self.check_file_1_p,
                       message=self.file_1_string_p,
                       delay=0.5, alpha=1, wraplength=200)
            # self.train_btn.configure(state="normal")
            self.remove_data_tree_p()
            self.canvas_frame_p.grid(row=1, column=0, padx=5, pady=5, sticky='NEWS')
            # self.canvas.grid(row=0, column=0, padx=5, pady=5, sticky='NEWS')
            self.canvas_p.delete(self.image_on_canvas_p)

            # place dnd image on the canvas
            self.image_p = Image.open("images/Gradebook_image3.png")
            self.original_image_p = self.image_p
            self.update_canvas()
            self.canvas_p.pack(expand=True, fill="both")

        else:
            print("Max number of datafiles reached")
        print("data stored")

    def update_canvas(self):
        if self.main_tabs.get() == "Train Model":

            self.canvas_width = self.canvas.winfo_reqwidth()
            self.canvas_height = self.canvas.winfo_reqheight()
            # self.original_image = self.image
            self.image = self.image.resize((self.canvas_width, self.canvas_height))
            self.photo = ImageTk.PhotoImage(image=self.image)
            self.image_on_canvas = self.canvas.create_image(int(self.canvas_width / 2), int(self.canvas_height / 2),
                                                            image=self.photo)
        else:
            self.canvas_width_p = self.canvas_p.winfo_reqwidth()
            self.canvas_height_p = self.canvas_p.winfo_reqheight()
            # self.original_image_p = self.image_p
            self.image_p = self.image_p.resize((self.canvas_width_p, self.canvas_height_p))
            self.photo_p = ImageTk.PhotoImage(image=self.image_p)
            self.image_on_canvas_p = self.canvas_p.create_image(int(self.canvas_width_p / 2),
                                                                int(self.canvas_height_p / 2),
                                                                image=self.photo_p)

    def get_path(self, event):
        # Get the filepath from drag n drop image. remove the curly brackets
        self.file_path = event.data
        if self.file_path[0] == '{' and self.file_path[-1] == '}':
            self.file_path = self.file_path[1:-1]
        self.filename = self.file_path.split("/")[-1]
        self.filename_var.set(self.filename)
        self.file_open()

    def get_path_p(self, event):
        # Get the filepath from drag n drop image. remove the curly brackets
        self.file_path_p = event.data
        if self.file_path_p[0] == '{' and self.file_path_p[-1] == '}':
            self.file_path_p = self.file_path_p[1:-1]
        self.filename_p = self.file_path_p.split("/")[-1]
        self.filename_var_p.set(self.filename_p)
        self.file_open()

    def test_function(self):
        print("test")

    def train_model(self):
        self.save_model.configure(state="normal")
        self.dynamic_right_frame.grid_remove()
        self.dynamic_right_frame.grid()

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

    def filter_data_p(self):
        # Create a copy of the column names
        columns = list(self.df_p.columns)

        # Iterate over the check_vars list and the column names together
        for check_var_p, column in zip(self.check_vars_p, columns):
            # If the checkbox is unchecked
            if not check_var_p.get():
                # Drop the column from the DataFrame
                self.df_p = self.df_p.drop(column, axis=1)

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
        self.update_idletasks()
        self.after_idle(lambda: self.load_data())

    def remove_zeros_p(self):

        # Calculate the threshold for number of zeros or '-' characters
        threshold = 0.7 * len(self.df_p.columns)

        # Define a function to apply to each row
        def check_row_p(row):
            # Count the number of zeros or '-' characters in the row
            zero_dash_count = sum((row == 0) | (row == '-'))
            # Return True if the count is less than the threshold (i.e., keep the row)
            # and False otherwise (i.e., remove the row)
            return zero_dash_count < threshold

        # Apply the function to each row and keep only the rows where the function returns True
        self.df_p = self.df_p[self.df_p.apply(check_row_p, axis=1)]
        print("zeros removed")
        self.update_idletasks()
        self.after_idle(lambda: self.load_data())

    def file_open(self):
        try:
            if self.main_tabs.get() == "Train Model":
                self.df = pd.read_excel(self.file_path)
                self.orig_df = self.df

            else:
                self.df_p = pd.read_excel(self.file_path_p)
                self.orig_df_p = self.df_p

            print("file opened")
            self.load_data()

        except ValueError:
            if self.main_tabs.get() == "Train Model":
                self.filename_var.set("Error", "The file you have chosen is invalid")
            else:
                self.filename_var_p.set("Error", "The file you have chosen is invalid")
            return

    def load_data(self):

        check_height = 50

        # place dnd image on the canvas
        if self.main_tabs.get() == "Train Model":
            self.image = Image.open("images/Loading.png")
            self.original_image = self.image
            self.canvas.delete(self.image_on_canvas)
            self.update_canvas()

            self.canvas.pack(expand=True, fill="both")
            # Create a new frame for the checkboxes
            self.checkboxes_frame_outer = ctk.CTkFrame(self.dynamic_right_frame, height=check_height, corner_radius=10,
                                                       fg_color=MAIN_FRAME_COLOR, border_width=0)

            self.checkboxes_frame_outer.grid(row=0, column=0, columnspan=1, padx=5, pady=0, sticky='news')

            # Add a Canvas to this frame
            self.checkboxes_canvas = Canvas(self.checkboxes_frame_outer, height=check_height, bg=LABEL_COLOR,
                                            borderwidth=0, bd=0, relief='ridge', highlightthickness=0)
            self.checkboxes_canvas.pack(side='left', fill='both', expand=True)

            # Create a new frame for the checkboxes and add it to the Canvas
            self.checkboxes_frame = ctk.CTkFrame(self.checkboxes_canvas, height=check_height, corner_radius=10,
                                                 fg_color=LABEL_COLOR, border_width=0)

            # adjusting 0,0 here might make it render quicker
            self.checkboxes_canvas.create_window((0, 0), window=self.checkboxes_frame, anchor='sw')
            self.clear_tree()
            # set up new tree view
            self.data_tree["column"] = list(self.df.columns)
            # replace below with empty string to remove headings line
            self.data_tree["show"] = ""
            # loop through column lists and add them to the tree view
            self.check_vars = []
            i = 0
            col_width = 120
            scaled_width = int(col_width * self.scale_factor)
            for column in self.data_tree["columns"]:
                check_var = BooleanVar()
                num_chars = 15
                short_col = column[:num_chars]
                self.check_vars.append(check_var)
                self.check_ind_frame = ctk.CTkFrame(self.checkboxes_frame, height=check_height, width=col_width)
                self.check_ind_frame.pack_propagate(False)
                self.check_ind_frame.grid_propagate(False)
                self.check_ind_frame.grid(row=0, column=i, padx=(0, 0), pady=0, ipadx=0, ipady=0, sticky="sew")
                check_button = ctk.CTkCheckBox(self.check_ind_frame, variable=check_var, command=self.on_check,
                                               text=short_col, width=15, height=15, font=("Inter", 10))

                # code below to enable tooltips
                CTkToolTip(check_button,
                           message=column,
                           delay=0.5, alpha=1, wraplength=200)

                check_button.pack(side='bottom', padx=(5, 0), expand=True, fill='both')
                check_button.pack_propagate(False)
                self.data_tree.column(column, stretch=False, width=scaled_width)
                self.data_tree.heading(column, command=lambda: "break", text=column)
                self.data_tree.bind('<Button-1>', self.do_nothing)
                i = i + 1

            # put the data in the treeview
            df_rows = self.df.to_numpy().tolist()
            for row in df_rows:
                self.data_tree.insert("", "end", values=row)

            self.update_idletasks()
            self.data_tree.grid(row=1, column=0, columnspan=1, padx=5, pady=5, ipadx=0, ipady=0, sticky='NEWS')
            self.data_tree.grid_propagate(False)
            # Add these lines where you create your Treeview widget
            self.yscrollbar = ctk.CTkScrollbar(self.dynamic_right_frame, command=self.data_tree.yview)
            self.data_tree.configure(yscrollcommand=self.yscrollbar.set)

            # Place the scrollbar next to the Treeview
            self.yscrollbar.grid(row=1, column=1, sticky='ns')

            self.x_scrollbar = ctk.CTkScrollbar(self.dynamic_right_frame, command=self.multi_scroll,
                                                orientation="horizontal")

            self.x_scrollbar.grid(row=2, column=0, sticky='ew')
            self.data_tree.configure(xscrollcommand=self.x_scrollbar.set)
            self.checkboxes_canvas.configure(xscrollcommand=self.x_scrollbar.set)

            # Update the scroll region of the Canvas
            self.checkboxes_canvas.configure(scrollregion=self.checkboxes_canvas.bbox('all'))

            self.after_idle(lambda: self.x_scrollbar.set(0, 0))
            self.after_idle(lambda: self.data_tree.xview_moveto(0))
            self.after_idle(lambda: self.checkboxes_canvas.xview_moveto(0))

            self.data_loaded = True
            self.update_idletasks()
            self.after_idle(lambda: self.enable_data_buttons())
            self.canvas_frame.grid_forget()
            self.after_idle(lambda: self.dynamic_right_frame.grid(row=1, column=0, padx=5, pady=5, sticky='NEWS'))
        else:
            self.image_p = Image.open("images/Loading.png")
            self.original_image_p = self.image_p
            self.canvas_p.delete(self.image_on_canvas_p)
            self.update_canvas()

            self.canvas_p.pack(expand=True, fill="both")
            # Create a new frame for the checkboxes
            # self.checkboxes_frame_outer = ctk.CTkFrame(self.train_frame_right, height =check_height, corner_radius=10, fg_color=MAIN_FRAME_COLOR,  border_width=0)
            self.checkboxes_frame_outer_p = ctk.CTkFrame(self.dynamic_right_frame_p, height=check_height,
                                                         corner_radius=10,
                                                         fg_color=MAIN_FRAME_COLOR, border_width=0)

            self.checkboxes_frame_outer_p.grid(row=0, column=0, columnspan=1, padx=5, pady=0, sticky='news')

            # Add a Canvas to this frame
            self.checkboxes_canvas_p = Canvas(self.checkboxes_frame_outer_p, height=check_height, bg=LABEL_COLOR,
                                              borderwidth=0, bd=0, relief='ridge', highlightthickness=0)
            self.checkboxes_canvas_p.pack(side='left', fill='both', expand=True)

            # Create a new frame for the checkboxes and add it to the Canvas
            self.checkboxes_frame_p = ctk.CTkFrame(self.checkboxes_canvas_p, height=check_height, corner_radius=10,
                                                   fg_color=LABEL_COLOR, border_width=0)

            # adjusting 0,0 here might make it render quicker
            self.checkboxes_canvas_p.create_window((0, 0), window=self.checkboxes_frame_p, anchor='sw')

            self.clear_tree_p()
            # set up new tree view
            self.data_tree_p["column"] = list(self.df_p.columns)
            # replace below with empty string to remove headings line
            self.data_tree_p["show"] = ""
            # loop through column lists and add them to the tree view
            self.check_vars_p = []
            i_p = 0
            col_width_p = 120
            scaled_width_p = int(col_width_p * self.scale_factor)
            for column in self.data_tree_p["columns"]:
                check_var_p = BooleanVar()
                num_chars = 15
                short_col = column[:num_chars]
                self.check_vars_p.append(check_var_p)
                self.check_ind_frame_p = ctk.CTkFrame(self.checkboxes_frame_p, height=check_height, width=col_width_p)
                self.check_ind_frame_p.pack_propagate(False)
                self.check_ind_frame_p.grid_propagate(False)
                self.check_ind_frame_p.grid(row=0, column=i_p, padx=(0, 0), pady=0, ipadx=0, ipady=0, sticky="sew")
                check_button_p = ctk.CTkCheckBox(self.check_ind_frame_p, variable=check_var_p, command=self.on_check_p,
                                                 text=short_col, width=15, height=15, font=("Inter", 10))

                # code below to enable tooltips
                CTkToolTip(check_button_p,
                           message=column,
                           delay=0.5, alpha=1, wraplength=200)

                check_button_p.pack(side='bottom', padx=(5, 0), expand=True, fill='both')
                check_button_p.pack_propagate(False)
                self.data_tree_p.column(column, stretch=False, width=scaled_width_p)
                self.data_tree_p.heading(column, command=lambda: "break", text=column)
                self.data_tree_p.bind('<Button-1>', self.do_nothing)
                i_p = i_p + 1

            # put the data in the treeview
            df_rows_p = self.df_p.to_numpy().tolist()
            for row in df_rows_p:
                self.data_tree_p.insert("", "end", values=row)

            self.update_idletasks()
            self.data_tree_p.grid(row=1, column=0, columnspan=1, padx=5, pady=5, ipadx=0, ipady=0, sticky='NEWS')
            self.data_tree_p.grid_propagate(False)
            # Add these lines where you create your Treeview widget
            self.yscrollbar_p = ctk.CTkScrollbar(self.dynamic_right_frame_p, command=self.data_tree.yview)
            self.data_tree_p.configure(yscrollcommand=self.yscrollbar_p.set)

            # Place the scrollbar next to the Treeview
            self.yscrollbar_p.grid(row=1, column=1, sticky='ns')

            self.x_scrollbar_p = ctk.CTkScrollbar(self.dynamic_right_frame_p, command=self.multi_scroll_p,
                                                  orientation="horizontal")

            self.x_scrollbar_p.grid(row=2, column=0, sticky='ew')
            self.data_tree_p.configure(xscrollcommand=self.x_scrollbar_p.set)
            self.checkboxes_canvas_p.configure(xscrollcommand=self.x_scrollbar_p.set)

            # Update the scroll region of the Canvas
            self.checkboxes_canvas_p.configure(scrollregion=self.checkboxes_canvas_p.bbox('all'))

            self.after_idle(lambda: self.x_scrollbar_p.set(0, 0))
            self.after_idle(lambda: self.data_tree_p.xview_moveto(0))
            self.after_idle(lambda: self.checkboxes_canvas_p.xview_moveto(0))

            self.data_loaded_p = True
            self.update_idletasks()
            self.after_idle(lambda: self.enable_data_buttons_p())
            self.canvas_frame_p.grid_forget()
            self.after_idle(lambda: self.dynamic_right_frame_p.grid(row=1, column=0, padx=5, pady=5, sticky='NEWS'))

    def on_resize(self, event):
        # Get the current scale factor
        # print("resized")
        self.get_scale_factor()
        if self.scale_factor != self.orig_scale:
            print("Scale Factor Changed")
            self.orig_scale = self.scale_factor

            # The following code will work to reopen and rescale the file for a new monitor DPI, but it is quite slow moving between monitors
            # current issue that the checkboxes stop scrolling fully
            if self.data_loaded == True:
                self.remove_data_tree()
                self.file_open()

        self.after_idle(self.update_image)

    def remove_data_tree(self):
        self.checkboxes_frame_outer.grid_forget()
        self.data_tree.grid_forget()
        self.x_scrollbar.grid_forget()
        self.yscrollbar.grid_forget()

    def update_image(self):
        if self.main_tabs.get() == "Train Model":

            self.canvas_width = self.canvas_frame.winfo_width()
            self.canvas_height = self.canvas_frame.winfo_height()

            self.image = self.original_image.resize((self.canvas_width, self.canvas_height))
            self.photo = ImageTk.PhotoImage(image=self.image)
            self.image_on_canvas = self.canvas.create_image(int(self.canvas_width / 2), int(self.canvas_height / 2),
                                                            image=self.photo)

        else:
            self.canvas_width_p = self.canvas_frame_p.winfo_width()
            self.canvas_height_p = self.canvas_frame_p.winfo_height()

            self.image_p = self.original_image_p.resize((self.canvas_width_p, self.canvas_height_p))
            self.photo_p = ImageTk.PhotoImage(image=self.image_p)
            self.image_on_canvas_p = self.canvas_p.create_image(int(self.canvas_width_p / 2),
                                                                int(self.canvas_height_p / 2),
                                                                image=self.photo_p)

    def get_scale_factor(self):
        winID = ctypes.windll.user32.GetForegroundWindow()

        MONITOR_DEFAULTTONULL = 0
        MONITOR_DEFAULTTOPRIMARY = 1
        MONITOR_DEFAULTTONEAREST = 2
        self.MDT_EFFECTIVE_DPI = 0

        monitorID = ctypes.windll.user32.MonitorFromWindow(winID, MONITOR_DEFAULTTONEAREST)
        dpiX = ctypes.c_uint()
        dpiY = ctypes.c_uint()
        ctypes.windll.shcore.GetDpiForMonitor(monitorID, self.MDT_EFFECTIVE_DPI, ctypes.byref(dpiX), ctypes.byref(dpiY))

        self.scale_factor = dpiX.value / 96
        # print("Current Scale Factor: ", self.scale_factor)

    def clear_tree(self):
        self.data_tree.delete(*self.data_tree.get_children())

    def clear_tree_p(self):
        self.data_tree_p.delete(*self.data_tree_p.get_children())

    def multi_scroll(self, *args):

        self.data_tree.xview(*args)
        self.checkboxes_canvas.xview(*args)

    def multi_scroll_p(self, *args):

        self.data_tree_p.xview(*args)
        self.checkboxes_canvas_p.xview(*args)
        print("scrolling")

    def do_nothing(event, self):
        return "break"

    def disable_data_buttons(self):
        self.filter_btn.configure(state="disabled")
        self.zeros_btn.configure(state="disabled")
        self.reset_btn.configure(state="disabled")
        self.store_btn.configure(state="disabled")

    def enable_data_buttons(self):
        self.filter_btn.configure(state="normal")
        self.zeros_btn.configure(state="normal")
        self.reset_btn.configure(state="normal")
        self.store_btn.configure(state="normal")

    def disable_data_buttons_p(self):
        self.filter_btn_p.configure(state="disabled")
        self.zeros_btn_p.configure(state="disabled")
        self.reset_btn_p.configure(state="disabled")
        self.store_btn_p.configure(state="disabled")

    def enable_data_buttons_p(self):
        self.filter_btn_p.configure(state="normal")
        self.zeros_btn_p.configure(state="normal")
        self.reset_btn_p.configure(state="normal")
        self.store_btn_p.configure(state="normal")

    def on_check(self):
        # This function will be called when a checkbox is clicked
        # You can add your own logic here
        for i, check_var in enumerate(self.check_vars):
            print(f"Checkbox {i} is {'checked' if check_var.get() else 'not checked'}")

    def on_check_p(self):
        # This function will be called when a checkbox is clicked
        # You can add your own logic here
        for i_p, check_var_p in enumerate(self.check_vars_p):
            print(f"Checkbox {i_p} is {'checked' if check_var_p.get() else 'not checked'}")


if __name__ == '__main__':
    app = MainApp()
    app.mainloop()