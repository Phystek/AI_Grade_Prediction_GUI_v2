#using version 5.22 of customtkinter
import customtkinter as ctk
from settings import *
from tkinterdnd2 import TkinterDnD, DND_ALL
from tkinter import filedialog, Canvas, StringVar, ttk,BooleanVar
import pandas as pd
from tkinter import font as tkFont
import openpyxl
from CTkToolTip import *

#importing pyautogui will cause the app to not rescale when moving between windows, poor resolution, but stable columns!
#import pyautogui
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
        self.train_frame_left = ctk.CTkFrame(master=self.train_frame, corner_radius=10, fg_color=LABEL_COLOR, width=100)
        self.train_frame_left.grid(row=0, column=0, padx=5, pady=5, sticky='NSEW')
        self.train_frame_left.rowconfigure(0, weight=1)
        self.train_frame_left.rowconfigure(1, weight=1)
        self.train_frame_left.rowconfigure(2, weight=10)
        self.train_frame_left.columnconfigure(0, weight=1)

        self.train_frame_right = ctk.CTkFrame(master=self.train_frame, corner_radius=10, fg_color=LABEL_COLOR, width=400)
        self.train_frame_right.grid(row=0, column=1, padx=5, pady=5, sticky='NSEW')
        self.train_frame_right.rowconfigure(0, weight=1)
        self.train_frame_right.rowconfigure(1, weight=1)
        self.train_frame_right.rowconfigure(2, weight=16)
        #self.train_frame_right.rowconfigure(3, weight=1)
        self.train_frame_right.columnconfigure(0, weight=1)
        self.train_frame_right.grid_propagate(False)

        #add widgets to left frame
        import_file_btn = (ctk.CTkButton(self.train_frame_left, text='Import Gradebook File', height=btn_height, width=80, command=self.file_dialog,
                                          font=("Inter", 14, "bold")))
        import_file_btn.grid(row=0, column=0, columnspan=1, padx=5, pady=10, sticky='NEW')
        self.filename_var = StringVar(value="Drag & drop file here")
        self.entryWidget = ctk.CTkEntry(master=self.train_frame_left,font=("Inter", 10, "italic"), textvariable=self.filename_var, height=22)
        self.entryWidget.grid(row=1, column=0, columnspan=1, padx=5, pady=3, ipadx=0, ipady=0, sticky='NEW')
        self.entryWidget.drop_target_register(DND_ALL)
        self.entryWidget.dnd_bind("<<Drop>>", self.get_path)

        #add widgets to right frame
        #excel viewer widget
        self.right_buttons_frame = ctk.CTkFrame(master=self.train_frame_right, corner_radius=10, fg_color=LABEL_COLOR)
        self.right_buttons_frame.grid(row=0, column=0, padx=5, pady=5, sticky='NSEW')
        self.right_buttons_frame.rowconfigure(0, weight=1)
        self.right_buttons_frame.columnconfigure(0, weight=1)
        self.right_buttons_frame.columnconfigure(1, weight=1)
        self.right_buttons_frame.columnconfigure(2, weight=1)

        filter_btn = (
            ctk.CTkButton(self.right_buttons_frame, text='Filter', height=btn_height, width=80, command=self.test_function,
                                    font=("Inter", 14, "bold")))
        filter_btn.grid(row=0, column=0, columnspan=1, padx=5, pady=10, sticky='NEW')
        zeros_btn = (
            ctk.CTkButton(self.right_buttons_frame, text='Remove Zeros', height=btn_height, width=80,
                                    font=("Inter", 14, "bold")))
        zeros_btn.grid(row=0, column=1, columnspan=1, padx=5, pady=10, sticky='NEW')
        reset_btn = (ctk.CTkButton(self.right_buttons_frame, text='Reset', height=btn_height, width=80,
                                    font=("Inter", 14, "bold")))
        reset_btn.grid(row=0, column=2, columnspan=1, padx=5, pady=10, sticky='NEW')


        self.data_tree=ttk.Treeview(master = self.train_frame_right)

        self.scaleFactor = ctypes.windll.shcore.GetScaleFactorForDevice(0) / 100
        #self.scaleFactor=1
        print(self.scaleFactor)

        self.checkboxes_frame_outer = None

    def file_dialog(self):
        #get path after an import button has been clicked ( no curly brackets with this method)
        self.file_path = filedialog.askopenfilename(filetypes=[("xlsx files", "*.xlsx")])
        # extract the filename from the full file path so it can be shown in the GUI
        self.filename = self.file_path.split("/")[-1]
        self.filename_var.set(self.filename)
        self.file_open()


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
        self.checkboxes_frame_outer.grid_forget()
        #self.data_tree.grid_forget()
        #self.x_scrollbar.grid_forget()
        #self.yscrollbar.grid_forget()
        i = 0
        col_width = 120
        check_height = 50
        scaled_width = int(col_width * self.scaleFactor)
        self.checkboxes_frame_outer = ctk.CTkFrame(self.train_frame_right, height =check_height, corner_radius=10, fg_color=MAIN_FRAME_COLOR,  border_width=0)
        self.checkboxes_frame_outer.grid(row=1, column=0, columnspan =1, padx=5, pady=0, sticky='news')

        # Add a Canvas to this frame
        self.checkboxes_canvas = Canvas(self.checkboxes_frame_outer, height =check_height, bg=LABEL_COLOR, borderwidth=0, bd=0, relief='ridge', highlightthickness=0)
        self.checkboxes_canvas.pack(side='left', fill='both', expand=True)

        # Create a new frame for the checkboxes and add it to the Canvas
        self.checkboxes_frame = ctk.CTkFrame(self.checkboxes_canvas, height =check_height, corner_radius=10, fg_color=LABEL_COLOR, border_width=0)

        #adjusting 0,0 here might make it render quicker
        self.checkboxes_canvas.create_window((0, 0), window=self.checkboxes_frame, anchor='sw')
        for column in self.data_tree["columns"]:
            check_var = BooleanVar()
            num_chars = 10
            short_col = column[:num_chars - 6]
            self.check_vars.append(check_var)

            self.check_ind_frame = ctk.CTkFrame(self.checkboxes_frame, height=check_height, width=col_width)
            self.check_ind_frame.grid(row=0, column=i, padx=(0, 0), pady=0, ipadx=0, ipady=0, sticky="sew")
            self.check_ind_frame.pack_propagate(False)
            self.check_ind_frame.grid_propagate(False)
            # font = ctkFont.Font(font=widget.cget("font"))
            # char_width= font.measure("0");
            check_button = ctk.CTkCheckBox(self.check_ind_frame, variable=check_var, command=self.on_check,
                                           text=short_col, width=15, height=15, font=("Inter", 10))
            # font = tkFont.Font(font=check_button.cget("font"))
            # char_width = font.measure("0")
            # check_button_width_pixels = char_width * num_chars

            # code below to enable tooltips
            CTkToolTip(check_button,
                       message=column,
                       delay=0.5, alpha=1, wraplength=200)

            # check_button.grid(row=0, column=0, sticky="EW")
            # check_button.grid(row=0, column=i,padx=(0,0), pady=0, sticky="")
            # check_button.pack(side='bottom', padx=(5,0), expand=True, fill='both')
            check_button.pack(side='bottom', padx=(5, 0), expand=True, fill='both')
            check_button.pack_propagate(False)

    def file_open(self):
        if self.filename:
            try:
                df=pd.read_excel(self.file_path)
                print("file opened")
                self.load_data(df)
            except ValueError:
                self.filename_var.set("Error", "The file you have chosen is invalid")
                return

    def load_data(self, df):
        #clear  old tree view
        self.clear_tree()
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
        self.data_tree["column"] = list(df.columns)
        #replace below with empty string to remove headings line
        self.data_tree["show"] = "headings"
        #loop through column lists and add them to the tree view
        self.check_vars = []
        i=0
        col_width=120
        scaled_width=int(col_width*self.scaleFactor)
        for column in self.data_tree["columns"]:

            check_var = BooleanVar()
            num_chars=10
            short_col= column[:num_chars-6]
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
        df_rows = df.to_numpy().tolist()
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
        # Update the scroll region of the Canvas
        #self.checkboxes_frame.update_idletasks()
        #self.checkboxes_canvas.configure(scrollregion=self.checkboxes_canvas.bbox('all'))
        print("scrollbar moved")


    def on_resize(self, event):
        # Get the current scale factor
        #current_scale_x, current_scale_y = get_display_scale()

        # If the scale factor has changed, the window has been moved to a different monitor
        #if current_scale_x != self.scale_x or current_scale_y != self.scale_y:
        #    print("Window has been moved to a different monitor")

           # Update the stored scale factor
        #self.scale_x, self.scale_y = current_scale_x, current_scale_y
        origScale=self.scaleFactor
        self.scaleFactor = ctypes.windll.shcore.GetScaleFactorForDevice(2) / 100
        #self.scaleFactor=1

        print(self.scaleFactor)

        winID = ctypes.windll.user32.GetForegroundWindow()
        print("This is your current window ID: ", winID)

        MONITOR_DEFAULTTONULL = 0
        MONITOR_DEFAULTTOPRIMARY = 1
        MONITOR_DEFAULTTONEAREST = 2
        self.MDT_EFFECTIVE_DPI=0

        monitorID = ctypes.windll.user32.MonitorFromWindow(winID, MONITOR_DEFAULTTONEAREST)
        print("This is your active monitor ID: ", monitorID)
        self.scaleFactor2 = ctypes.windll.shcore.GetScaleFactorForDevice(2) / 100
        print(self.scaleFactor2)
        MONITORENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_int, ctypes.c_ulong, ctypes.c_ulong, ctypes.POINTER(RECT),
                                             ctypes.c_double)
        # Enumerate all monitors
        #ctypes.windll.user32.EnumDisplayMonitors(0, 0, MONITORENUMPROC(self.monitor_enum_proc), 0)
        hMonitor=monitorID
        #check if the monitor id has changed
        dpiX = ctypes.c_uint()
        dpiY = ctypes.c_uint()
        ctypes.windll.shcore.GetDpiForMonitor(hMonitor, self.MDT_EFFECTIVE_DPI, ctypes.byref(dpiX), ctypes.byref(dpiY))

        self.currentdpi  = dpiX.value / 96

        print("Current DPI: ", self.currentdpi)


        # Define callback function
    def monitor_enum_proc(self, hMonitor, hdcMonitor, lprcMonitor, dwData):
        # Get the DPI for the monitor
        dpiX = ctypes.c_uint()
        dpiY = ctypes.c_uint()
        ctypes.windll.shcore.GetDpiForMonitor(hMonitor, self.MDT_EFFECTIVE_DPI, ctypes.byref(dpiX), ctypes.byref(dpiY))
        scaleFactorX = dpiX.value / 96
        scaleFactorY = dpiY.value / 96
        print(f"Monitor ID: {hMonitor}, Scale Factor: {scaleFactorX}, {scaleFactorY}")
        return 1  # Continue enumeration

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



def get_display_scale():
    screen_size = pyautogui.size()
    monitor = get_monitors()[0]  # Get the first monitor

    return screen_size[0] / monitor.width, screen_size[1] / monitor.height


class RECT(ctypes.Structure):
    _fields_ = [("left", ctypes.c_long),
                ("top", ctypes.c_long),
                ("right", ctypes.c_long),
                ("bottom", ctypes.c_long)]







if __name__ == '__main__':
    app = MainApp()
    app.mainloop()