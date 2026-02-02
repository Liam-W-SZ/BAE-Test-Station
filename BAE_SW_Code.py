import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
from CTkMessagebox import CTkMessagebox
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import json
import os
import glob
import numpy as np
from tv_tools import root, outputJSON_local, test, outputEXCEL_local, outputJSON, outputEXCEL
from matplotlib.backend_bases import MouseButton
import time
import io

# Optional SharePoint support
try:
    from msal import ConfidentialClientApplication
    import requests
    SHAREPOINT_AVAILABLE = True
except ImportError:
    SHAREPOINT_AVAILABLE = False
    print("SharePoint libraries not available. Install with: pip install msal requests")

# Set appearance and color theme
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class ScrollableOptionMenu(ctk.CTkOptionMenu):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.configure(height=10)

# Main application class for BAE Test Station.
class BAETestApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        # Window setup
        self.title("BAE Test Station")
        self.geometry("1000x850")
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=2)
        self.sidebar = ctk.CTkFrame(self, width=400)
        self.sidebar.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        self.main_content1 = ctk.CTkFrame(self)
        self.main_content1.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")
        self.main_content1.grid_columnconfigure(0, weight=1)
        self.current_operator = None
        self.data = None
        self.all_file_names = []
        self.setup_sidebar()
        self.setup_main_content1()
        self.populate_file_dropdown()

        # Load login config
        with open('BAE_Login.json', 'r') as f:
            self.configLogin = json.load(f)

        # Deafault Data
        self.ts = "ts"
        self.test_desc = "Battery Aging Evaluation Test"
        self.test_jig = "NA"
        self.device_id = "NA"
        self.serialnr = "NA"
        self.jobnr = "NA"
        self.prefix = "NA"
        self.procedure = "BAE Test"
        self.productGroup = "Batteries"
        self.supplierserial = "NA"
        self.lowerlevelID = "NA"
        self.tests = []
        self.errors = []
        self.result = "Undetermined"
        self.alarms = []
        self.concession_pass_points = []
        self.fail_point = []
        self.Jsonalarms_errors = []
        self.protocol("WM_DELETE_WINDOW", self.quit)
        self.current_canvas = None
        self.set_panels_state("disabled")

        # SharePoint config
        try:
            with open('BAE_SharePoint_Config.json', 'r') as f:
                sharepoint_config_data = json.load(f)
                self.sharepoint_config = sharepoint_config_data["SharePoint_Config"]
        except FileNotFoundError:
            print("[WARNING] SharePoint config file not found, using default values")
            self.sharepoint_config = {
                "sharepoint_url": "https://stage0.sharepoint.com/sites/Assurance-Quality",
                "excel_file_url": "/sites/Assurance-Quality/Build Histories/01. Build Histories/TDV Line/Marconi Connect.xlsx",
                "folder_url": "/sites/Assurance-Quality/Build Histories/01. Build Histories/TDV Line",
                "username": "productionchecks@stagezero.co.za",
                "password": "Pr0ductCh3ck2*"
            }
        self.sharepoint_files = []

    # Enable / disable all relevant panels, buttons etc
    def set_panels_state(self, state):
        self.start_button.configure(state=state)
        self.reset_button.configure(state=state)
        self.help_button.configure(state="normal")
        self.quit_button.configure(state=state)
        self.file_dropdown_menu.configure(state=state)
        self.DC_checkbox.configure(state=state)
        self.refresh_button.configure(state=state)
        self.local_radio.configure(state=state)
        self.sharepoint_radio.configure(state=state)
        self.dropdown_menu1.configure(state=state)
        self.display_button1.configure(state=state)
        self.clear_button1.configure(state=state)

    # Store test data and handle result file output and uploads
    def store_test_data(self):
        global ts, tester, testdesc, testjig, deviceid, serialnr, jobnr, prefix, procedure, productGroup, supplierserial, lowerlevelID, tests, errors, result, TestFile1

        ts = int(time.time())  # epoch time in seconds
        tester = self.current_operator if self.current_operator else "Unknown"  # Current operator or "Unknown"
        testdesc = self.test_desc  # default
        testjig = self.test_jig  # Default
        deviceid = self.device_id  # Get
        serialnr = self.serialnr  # Get
        jobnr = self.jobnr  # Assuming this is a constant or incremented elsewhere
        prefix = self.prefix  # Default prefix
        procedure = self.procedure  # Default procedure
        productGroup = self.productGroup  # Default product group
        supplierserial = self.supplierserial
        lowerlevelID = self.lowerlevelID  # Default lower level ID
        tests = self.tests  # Default tests
        errors = self.Jsonalarms_errors  # Default errors
        result = self.result  # Default result, e.g., "Passed"

        # Create TestFile1 object
        TestFile1 = root(
            ts, tester, testdesc, testjig, deviceid, prefix, jobnr, serialnr,
            procedure, productGroup, supplierserial, lowerlevelID, tests, errors, result
        )

        # Open the BAE_File_Paths config file to get the Result_Folder_Path
        with open("BAE_File_Paths.json", "r") as file:
            config_data = json.load(file)
        
        # JSON
        result_folder_path = config_data["Json_Result_Folder_Path"]
        result_filename = f"{deviceid}_{ts}_result_{result}.json"

        # Excel
        excel_output_folder_path = config_data["Excel_Result_Folder_Path"]  # New folder
        original_excel_path = config_data["Excel_TestData_Folder_Path"]  # Old folder
        
        # Use the loaded filename for Excel operations
        excel_filename = getattr(self, 'loaded_filename', deviceid)
        excel_base_name = os.path.splitext(excel_filename)[0]
        excel_extension = os.path.splitext(excel_filename)[1]  # Get the original extension

        if result == "Solid_Pass":
            result_filepath = os.path.join(result_folder_path, "Pass", result_filename)
            outputJSON_local(TestFile1, result_filepath)  
            excel_result_path = os.path.join(excel_output_folder_path, "Pass", f"{excel_base_name}_Pass{excel_extension}")
            outputEXCEL_local((excel_output_folder_path + "//Pass"), original_excel_path, excel_filename, "_Pass")  

            # FTP Server Upload Json - Pass
            result_Json_ftp_folder = 'BAE_Test_Folder/BAE_Json_Results/Pass'
            outputJSON(TestFile1, result_filepath, result_filename, result_Json_ftp_folder)

            # FTP Server Upload Excel - Pass
            result_Excel_ftp_folder = 'BAE_Test_Folder/BAE_Excel_Results/Pass'
            outputEXCEL((excel_output_folder_path + "//Pass"), f"{excel_base_name}_Pass", result_Excel_ftp_folder, excel_extension)

            # SharePoint Upload - Pass
            self.upload_test_results_to_sharepoint("Solid_Pass", result_filepath, excel_result_path, deviceid)

            # Delete source file after successful processing
            self.delete_source_file()

        if result == "Concession_Pass":
            result_filepath = os.path.join(result_folder_path, "Concession_Pass", result_filename)
            outputJSON_local(TestFile1, result_filepath) 
            excel_result_path = os.path.join(excel_output_folder_path, "Concession_Pass", f"{excel_base_name}_Concession_Pass{excel_extension}")
            outputEXCEL_local((excel_output_folder_path + "//Concession_Pass"), original_excel_path, excel_filename, "_Concession_Pass")

            # FTP Server Upload Json - Concession Pass
            result_Json_ftp_folder = 'BAE_Test_Folder/BAE_Json_Results/Concession_Pass'
            outputJSON(TestFile1, result_filepath, result_filename, result_Json_ftp_folder)

            # FTP Server Upload Excel - Concession_Pass
            result_Excel_ftp_folder = 'BAE_Test_Folder/BAE_Excel_Results/Concession_Pass'
            outputEXCEL((excel_output_folder_path + "//Concession_Pass"), f"{excel_base_name}_Concession_Pass", result_Excel_ftp_folder, excel_extension)

            # SharePoint Upload - Concession Pass
            self.upload_test_results_to_sharepoint("Concession_Pass", result_filepath, excel_result_path, deviceid)

            # Delete source file after successful processing
            self.delete_source_file()

        if result == "Fail":
            result_filepath = os.path.join(result_folder_path, "Fail", result_filename)
            outputJSON_local(TestFile1, result_filepath)  # Write the larger JSON file
            excel_result_path = os.path.join(excel_output_folder_path, "Fail", f"{excel_base_name}_Fail{excel_extension}")
            outputEXCEL_local((excel_output_folder_path + "//Fail"), original_excel_path, excel_filename, "_Fail")

            # FTP Server Upload Json - Fail
            result_Json_ftp_folder = 'BAE_Test_Folder/BAE_Json_Results/Fail'
            outputJSON(TestFile1, result_filepath, result_filename, result_Json_ftp_folder)

            # FTP Server Upload Excel - Fail
            result_Excel_ftp_folder = 'BAE_Test_Folder/BAE_Excel_Results/Fail'
            outputEXCEL((excel_output_folder_path + "//Fail"), f"{excel_base_name}_Fail", result_Excel_ftp_folder, excel_extension)

            # SharePoint Upload - Fail
            self.upload_test_results_to_sharepoint("Fail", result_filepath, excel_result_path, deviceid)

            # Delete source file after successful processing
            self.delete_source_file()

    # Delete source file after processing           
    def delete_source_file(self):
        try:
            if not hasattr(self, 'loaded_filename') or not self.loaded_filename:
                print("[WARNING] No loaded filename found, cannot delete source file")
                return False
            
            source = self.file_source_var.get()
            
            if source == "Local":
                # Delete local file from Aging_Data folder
                with open("BAE_File_Paths.json", "r") as file:
                    config_data = json.load(file)
                
                aging_data_folder = config_data["Excel_TestData_Folder_Path"]
                source_file_path = os.path.join(aging_data_folder, self.loaded_filename)
                
                if os.path.exists(source_file_path):
                    os.remove(source_file_path)
                    print(f"[SUCCESS] Local source file deleted: {source_file_path}")
                    return True
                else:
                    print(f"[WARNING] Local source file not found: {source_file_path}")
                    return False
                    
            elif source == "SharePoint":
                # Delete SharePoint file from untested data folder
                success = self.delete_sharepoint_file(self.loaded_filename)
                if success:
                    print(f"[SUCCESS] SharePoint source file deleted: {self.loaded_filename}")
                else:
                    print(f"[ERROR] Failed to delete SharePoint source file: {self.loaded_filename}")
                return success
            else:
                print(f"[WARNING] Unknown file source: {source}")
                return False
                
        except Exception as e:
            print(f"[ERROR] Error deleting source file: {e}")
            return False

    def setup_sidebar(self):
        # Title
        self.logo_label = ctk.CTkLabel(
            self.sidebar, 
            text="BAE Test Station",
            font=ctk.CTkFont(size=20, weight="bold")
        )
        self.logo_label.pack(pady=20, padx=20)

        #Login
        self.auth_frame = ctk.CTkFrame(self.sidebar)
        self.auth_frame.pack(pady=20, fill="x", padx=20)

        self.lock_button = ctk.CTkButton(
            self.auth_frame,
            state="normal",
            text="",
            width=20,
            height=20,
            fg_color="red",  # Start with red (logged out)
            hover=False,
            command=self.toggle_login
        )
        self.lock_button.pack(side="left", padx=5)
        
        self.operator_label = ctk.CTkLabel(
            self.auth_frame,
            text="Not logged in",
            font=ctk.CTkFont(size=12)
        )
        self.operator_label.pack(side="left", padx=5)
        
        # Status indicator
        self.status_frame = ctk.CTkFrame(self.sidebar)
        self.status_frame.pack(pady=20, fill="x", padx=20)

        self.status_label = ctk.CTkLabel(
            self.status_frame,
            text="Status:",
            font=ctk.CTkFont(size=14)
        )
        self.status_label.pack(side="left", padx=5)

        self.status_indicator = ctk.CTkLabel(
            self.status_frame,
            text="Not logged in",
            font=ctk.CTkFont(size=14),
            text_color="yellow"
        )
        self.status_indicator.pack(side="left", padx=5)

        #Confirm if old data was cleared before Aging occurred
        self.serial_label_DC = ctk.CTkLabel(
            self.sidebar,
            text="Old Aging Data Cleared:",
            font=ctk.CTkFont(size=14)
        )
        self.serial_label_DC.pack(pady=5, padx=20)
        #self.logo_label.pack(pady=20, padx=20)

        #Inspection Tickbox
        self.DC_checkbox = ctk.CTkCheckBox(
            self.sidebar,
            text="Data was Cleared",
            state="disabled",  # Initially disabled
            #command=self.on_checkbox_change,
        )
        self.DC_checkbox.pack(pady=(0,20))

        #Label for file source selection
        self.file_source_label = ctk.CTkLabel(
            self.sidebar,
            text="Select File Source:",
            font=ctk.CTkFont(size=14)
        )
        self.file_source_label.pack(pady=5, padx=20)

        # File source selection (Local vs SharePoint)
        self.file_source_var = tk.StringVar(value="Local")
        self.source_frame = ctk.CTkFrame(self.sidebar)
        self.source_frame.pack(pady=(0, 10), padx=20, fill="x")

        self.local_radio = ctk.CTkRadioButton(
            self.source_frame,
            text="Local Files",
            variable=self.file_source_var,
            value="Local",
            command=self.on_source_changed
        )
        self.local_radio.pack(side="left", padx=5)

        self.sharepoint_radio = ctk.CTkRadioButton(
            self.source_frame,
            text="SharePoint",
            variable=self.file_source_var,
            value="SharePoint",
            command=self.on_source_changed
        )
        self.sharepoint_radio.pack(side="left", padx=5)

        #Label for excel file selection
        self.serial_label_DDM = ctk.CTkLabel(
            self.sidebar,
            text="Select Excel File:",
            font=ctk.CTkFont(size=14)
        )
        self.serial_label_DDM.pack(pady=5, padx=20)

        # Search entry above the dropdown
        self.file_search_var = tk.StringVar()
        self.file_search_entry = ctk.CTkEntry(
            self.sidebar,
            textvariable=self.file_search_var,
            width=200,
            placeholder_text="Search Excel Files"
        )
        self.file_search_entry.pack(pady=(0, 5))
        self.file_search_var.trace_add("write", self.filter_file_dropdown)

        # Refresh button for SharePoint and local files
        self.refresh_button = ctk.CTkButton(
            self.sidebar,
            text="Refresh Files",
            command=self.refresh_files,
            width=50,
            height=30,
            fg_color="#800080",
            hover_color="#9932CC"
        )
        self.refresh_button.pack(pady=(0, 5))

        # Dropdown for file selection
        self.file_dropdown_menu = ScrollableOptionMenu(
            self.sidebar,
            values=[],
            width=250,   
            font=("Arial", 16),         
            command=self.on_file_selected,  
        )
        self.file_dropdown_menu.pack(pady=(0,20))
        self.file_dropdown_menu.configure(
            fg_color="#800080",  
            dropdown_fg_color="#800080",  
            button_color="#800080",  
            button_hover_color="#9932CC"  
        )
        self.file_dropdown_menu.set("Excel Files")  # Set default text

        # Buttons
        # Start Test Button
        self.start_button = ctk.CTkButton(
            self.sidebar,
            text="Start Test",
            command=self.start_test,
            width=200,
            height=40,
            state="disabled",  
            fg_color="#800080",  
            hover_color="#9932CC"  
        )
        self.start_button.pack(pady=10)

        # Reset Button
        self.reset_button = ctk.CTkButton(
            self.sidebar,
            text="Reset",
            command=self.reset_gui,
            width=200,
            height=40,
            fg_color="transparent",
            border_width=2,
            state="normal"  # Initially enabled
        )
        self.reset_button.pack(pady=10)

        # Help Button
        self.help_button = ctk.CTkButton(
            self.sidebar,
            text="Help",
            command=self.open_help_file,
            width=200,
            height=40,
            fg_color="#800080", 
            hover_color="#9932CC"  
        )
        self.help_button.pack(pady=10)

        # Quit Button - not really necessary, but awe
        self.quit_button = ctk.CTkButton(
            self.sidebar,
            text="Quit",
            command=self.quit,
            width=200,
            height=40,
            state="normal",  
            fg_color="#800080",  
            hover_color="#9932CC"  
        )
        self.quit_button.pack(pady=10)

    def open_help_file(self):
        help_file_path = os.path.join(os.path.dirname(__file__), "Help.txt")
        try:
            os.startfile(help_file_path)  # Open the help file with the default text editor
        except FileNotFoundError:
            CTkMessagebox(
                title="Error",
                message="Help file not found. Please ensure 'Help.txt' exists in the program directory.",
                icon="cancel"
            )
    # Setup area for results and graphs
    def setup_main_content1(self):
        # Configure main content grid
        self.main_content1.grid_columnconfigure(0, weight=1)

        # Results display
        self.result_text1 = ctk.CTkTextbox(
            self.main_content1,            
            width=300,
            height=300,
            font=ctk.CTkFont(size=12)
        )
        self.result_text1.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="nsew")
        #self.result_text1.insert("end", "Battery Criteria🔋\n")
        
        # Graph controls
        self.graph_controls_frame1 = ctk.CTkFrame(self.main_content1)
        self.graph_controls_frame1.grid(row=2, column=0, padx=20, pady=(10, 20), sticky="nsew")

        #Graph drop down menu
        self.dropdown_menu1 = ctk.CTkOptionMenu(
            self.graph_controls_frame1,
            #text = "Select File",
            values=["Battery voltage(V)", "Remaining capacity(Ah)", "Delta Cell Voltages"],
            width=150
        )
        self.dropdown_menu1.pack(side="left", padx=10)
        self.dropdown_menu1.configure(
            fg_color="#800080",  
            dropdown_fg_color="#800080",  
            button_color="#800080",  
            button_hover_color="#9932CC"  
        )

        #Graph display button
        self.display_button1 = ctk.CTkButton(
            self.graph_controls_frame1,
            text="Display Graph",
            command=self.display_graph1,
            width=150,
            fg_color="#800080",  
            hover_color="#9932CC"  
        )
        self.display_button1.pack(side="left", padx=10)      

        #Graph clear button
        self.clear_button1 = ctk.CTkButton(
            self.graph_controls_frame1,
            text="Clear Graph",
            command=self.clear_graph1,
            width=150,
            fg_color="#800080",  
            hover_color="#9932CC"  
        )
        self.clear_button1.pack(side="left", padx=10)


        # Graph display area
        self.graph_frame1 = ctk.CTkFrame(self.main_content1, height=300) #heigh was 350
        self.graph_frame1.grid(row=3, column=0, padx=20, pady=(10, 20), sticky="nsew")

    # Populate the file dropdown with local Excel/CSV files
    def populate_file_dropdown(self):
        # Open the BAE_File_Paths config file
        with open("BAE_File_Paths.json", "r") as file:
            config_data = json.load(file)
        
        # Use the Excel_TestData_Folder_Path from the config file
        folder_path = config_data["Excel_TestData_Folder_Path"]
        
        # Include both .xlsx, .xls, and .csv files
        excel_files = glob.glob(os.path.join(folder_path, "*.xlsx")) + \
                      glob.glob(os.path.join(folder_path, "*.xls")) + \
                      glob.glob(os.path.join(folder_path, "*.csv"))
        
        file_names = [os.path.basename(file) for file in excel_files]
        self.all_file_names = file_names  # Store all file names for filtering
        if file_names:
            self.file_dropdown_menu.configure(values=file_names)
            self.file_dropdown_menu.set("Excel Files")
        else:
            self.file_dropdown_menu.configure(values=["No More Excel Files"])
            self.file_dropdown_menu.set("No More Excel Files")
        self.start_button.configure(state="disabled")  

    # Handle file source change 
    def on_source_changed(self):
        source = self.file_source_var.get()
        if source == "Local":
            self.populate_file_dropdown()
            self.refresh_button.configure(text="Refresh Local")
        else:  # SharePoint
            # Reload SharePoint configuration when switching to SharePoint
            self.reload_sharepoint_config()
            self.refresh_files()
            self.refresh_button.configure(text="Refresh SharePoint")

    # Refresh file list based on selected source
    def refresh_files(self):
        source = self.file_source_var.get()
        
        if source == "Local":
            self.populate_file_dropdown()
        else:  # SharePoint
            if SHAREPOINT_AVAILABLE:
                # Reload SharePoint configuration first
                self.reload_sharepoint_config()
                
                # Show loading message
                self.file_dropdown_menu.configure(values=["Loading SharePoint files..."])
                self.file_dropdown_menu.set("Loading SharePoint files...")
                self.update()  # GUI update
                
                sharepoint_files = self.get_sharepoint_files()
                self.sharepoint_files = sharepoint_files
                self.all_file_names = sharepoint_files
                
                if sharepoint_files:
                    self.file_dropdown_menu.configure(values=sharepoint_files)
                    self.file_dropdown_menu.set("SharePoint Files")
                else:
                    self.file_dropdown_menu.configure(values=["No SharePoint Files"])
                    self.file_dropdown_menu.set("No SharePoint Files")
            else:
                self.file_dropdown_menu.configure(values=["SharePoint Not Available"])
                self.file_dropdown_menu.set("SharePoint Not Available")
        
        self.start_button.configure(state="disabled")  

    # Filter the file dropdown based on search input
    def filter_file_dropdown(self, *args):
        search_text = self.file_search_var.get().lower()
        if not search_text:
            filtered = self.all_file_names
        else:
            filtered = [f for f in self.all_file_names if search_text in f.lower()]
        if filtered:
            self.file_dropdown_menu.configure(values=filtered)
            self.file_dropdown_menu.set("Excel Files")
        else:
            self.file_dropdown_menu.configure(values=["No files found"])
            self.file_dropdown_menu.set("No files found")

    # Handle file selection from dropdown
    def on_file_selected(self, selected_file):
        # Only enable Start if a real file is selected
        invalid_selections = [
            "Excel Files", "Select Excel File", "No More Excel Files",
            "SharePoint Files", "No SharePoint Files", "SharePoint Not Available",
            "Loading SharePoint files...", "No files found"
        ]
        
        if selected_file and selected_file not in invalid_selections:
            self.start_button.configure(state="normal")
            self.load_selected_file(selected_file)
        else:
            self.start_button.configure(state="disabled")

    # Load the selected file (local / SharePoint)
    def load_selected_file(self, selected_file):
        source = self.file_source_var.get()
        
        if source == "Local":
            with open("BAE_File_Paths.json", "r") as file:
                config_data = json.load(file)
            
            # Use the Excel_TestData_Folder_Path from the config file
            folder_path = config_data["Excel_TestData_Folder_Path"]
            file_path = os.path.join(folder_path, selected_file)
            
            # Load the file based on its extension
            try:
                if selected_file.endswith(".csv"):
                    # Enhanced CSV loading with better delimiter detection
                    print(f"[INFO] Loading CSV file: {selected_file}")
                    
                    # Try different CSV parsing strategies
                    csv_strategies = [
                        # Strategy 1: Semicolon delimiter (European CSV / CSV UTF-8)
                        {
                            'sep': ';',
                            'encoding': 'utf-8',
                            'on_bad_lines': 'skip',
                            'quoting': 1,
                            'skipinitialspace': True
                        },
                        # Strategy 2: Semicolon with UTF-8 BOM (CSV UTF-8 from Excel)
                        {
                            'sep': ';',
                            'encoding': 'utf-8-sig',
                            'on_bad_lines': 'skip',
                            'quoting': 1,
                            'skipinitialspace': True
                        },
                        # Strategy 3: Standard CSV with comma delimiter
                        {
                            'sep': ',',
                            'encoding': 'utf-8',
                            'on_bad_lines': 'skip',
                            'quoting': 1,
                            'skipinitialspace': True
                        },
                        # Strategy 4: Comma with UTF-8 BOM
                        {
                            'sep': ',',
                            'encoding': 'utf-8-sig',
                            'on_bad_lines': 'skip',
                            'quoting': 1,
                            'skipinitialspace': True
                        },
                        # Strategy 5: Handle extra columns by limiting to first 30
                        {
                            'sep': ',',
                            'encoding': 'utf-8',
                            'on_bad_lines': 'skip',
                            'quoting': 1,
                            'skipinitialspace': True,
                            'usecols': range(30)
                        },
                        # Strategy 6: Different encoding with semicolon
                        {
                            'sep': ';',
                            'encoding': 'latin-1',
                            'on_bad_lines': 'skip',
                            'quoting': 1,
                            'skipinitialspace': True
                        },
                        # Strategy 7: Very permissive settings
                        {
                            'sep': ',',
                            'encoding': 'utf-8',
                            'on_bad_lines': 'skip',
                            'quoting': 3,
                            'skipinitialspace': True,
                            'engine': 'python'
                        }
                    ]
                    
                    self.data = None
                    strategy_used = None
                    
                    for i, strategy in enumerate(csv_strategies, 1):
                        try:
                            print(f"[INFO] Trying CSV strategy {i}: {strategy}")
                            self.data = pd.read_csv(file_path, **strategy)
                            
                            # Check if we got proper column separation (more than 1 column)
                            if len(self.data.columns) > 1:
                                strategy_used = i
                                print(f"[SUCCESS] CSV loaded with strategy {i}")
                                print(f"[INFO] Columns detected: {len(self.data.columns)}")
                                break
                            else:
                                print(f"[WARNING] Strategy {i} only detected {len(self.data.columns)} column(s) - likely wrong delimiter")
                                # Don't break here, try next strategy
                                continue
                                
                        except Exception as strategy_error:
                            print(f"[WARNING] Strategy {i} failed: {strategy_error}")
                            continue
                    
                    # If no strategy worked or we only got 1 column, try auto-detection
                    if self.data is None or len(self.data.columns) <= 1:
                        print("[INFO] Trying automatic delimiter detection...")
                        try:
                            import csv
                            with open(file_path, 'r', encoding='utf-8') as f:
                                sample = f.read(1024)
                                sniffer = csv.Sniffer()
                                delimiter = sniffer.sniff(sample).delimiter
                                print(f"[INFO] Auto-detected delimiter: '{delimiter}'")
                                
                                self.data = pd.read_csv(file_path, sep=delimiter, encoding='utf-8')
                                if len(self.data.columns) > 1:
                                    strategy_used = "Auto-detection"
                                    print(f"[SUCCESS] CSV loaded with auto-detection")
                                else:
                                    raise Exception("Auto-detection also failed to separate columns properly")
                        except Exception as auto_error:
                            print(f"[ERROR] Auto-detection failed: {auto_error}")
                            raise Exception("All CSV loading strategies failed. The file may have formatting issues.")
                    
                    # Final validation
                    if self.data is None:
                        raise Exception("All CSV loading strategies failed. The file may have formatting issues.")
                    
                    # Validate the loaded data
                    print(f"[INFO] CSV loaded successfully with {len(self.data)} rows and {len(self.data.columns)} columns")
                    print(f"[INFO] Strategy used: {strategy_used}")
                    print(f"[INFO] Column names: {list(self.data.columns)}")
                    
                    # Check for required columns
                    required_columns = ['Battery voltage(V)', 'Remaining capacity(Ah)', 'Alarm']
                    missing_columns = [col for col in required_columns if col not in self.data.columns]
                    
                    if missing_columns:
                        # Try to find similar column names
                        available_columns = list(self.data.columns)
                        print(f"[WARNING] Missing columns: {missing_columns}")
                        print(f"[INFO] Available columns: {available_columns}")
                        
                        # Enhanced column mapping
                        column_mapping = {}
                        for required_col in required_columns:
                            for available_col in available_columns:
                                if required_col.lower() in available_col.lower() or available_col.lower() in required_col.lower():
                                    column_mapping[available_col] = required_col
                                    break
                        
                        if column_mapping:
                            print(f"[INFO] Found similar columns: {column_mapping}")
                            self.data = self.data.rename(columns=column_mapping)
                            print("[SUCCESS] Column names mapped successfully")
                        else:
                            raise Exception(f"Required columns not found: {missing_columns}")
                
                else:  # Excel files (.xls, .xlsx)
                    print(f"[INFO] Loading Excel file: {selected_file}")
                    self.data = pd.read_excel(file_path)
                
                print(f"[SUCCESS] Local file loaded: {selected_file}")
                
            except Exception as e:
                error_msg = f"Failed to load local file: {str(e)}"
                print(f"[ERROR] {error_msg}")
                self.data = None
                return
        
        else:  # SharePoint
            # Load from SharePoint
            self.data = self.load_sharepoint_file(selected_file)
            if self.data is None:
                return
        
            file_ext = os.path.splitext(selected_file)[1].lower()

            # Load the BAE_File_Paths config for the test data folder
            with open("BAE_File_Paths.json", "r") as file:
                config_data = json.load(file)
            local_testdata_folder = config_data["Excel_TestData_Folder_Path"]
            local_file_path = os.path.join(local_testdata_folder, selected_file)
            try:
                if file_ext == ".csv":
                    self.data.to_csv(local_file_path, index=False)
                else:
                    self.data.to_excel(local_file_path, index=False)
                print(f"[INFO] Saved SharePoint file locally as: {local_file_path}")
            except Exception as save_err:
                print(f"[WARNING] Could not save SharePoint file locally: {save_err}")

        self.loaded_filename = selected_file
        self.device_id = os.path.splitext(selected_file)[0]

        if self.data is not None:
            # Check if required columns exist before cleaning
            required_columns = ['Battery voltage(V)', 'Remaining capacity(Ah)']
            existing_columns = [col for col in required_columns if col in self.data.columns]
            
            if existing_columns:
                print(f"[INFO] Cleaning data for columns: {existing_columns}")
                
                # Clean and convert voltage data first
                if 'Battery voltage(V)' in self.data.columns:
                    try:
                        # Convert string values with comma decimal separators to numeric
                        self.data['Battery voltage(V)'] = self.data['Battery voltage(V)'].astype(str)
                        self.data['Battery voltage(V)'] = self.data['Battery voltage(V)'].str.replace(',', '.')
                        self.data['Battery voltage(V)'] = pd.to_numeric(self.data['Battery voltage(V)'], errors='coerce')
                        
                        # Remove rows with invalid voltage data
                        self.data = self.data.dropna(subset=['Battery voltage(V)'])
                        self.data = self.data[self.data['Battery voltage(V)'] > 0]
                        print(f"[INFO] Cleaned voltage data: {len(self.data)} rows remaining")
                    except Exception as e:
                        print(f"[ERROR] Failed to clean voltage data: {e}")
                
                # Clean and convert capacity data
                if 'Remaining capacity(Ah)' in self.data.columns:
                    try:
                        # Convert string values with comma decimal separators to numeric
                        self.data['Remaining capacity(Ah)'] = self.data['Remaining capacity(Ah)'].astype(str)
                        self.data['Remaining capacity(Ah)'] = self.data['Remaining capacity(Ah)'].str.replace(',', '.')
                        self.data['Remaining capacity(Ah)'] = pd.to_numeric(self.data['Remaining capacity(Ah)'], errors='coerce')
                        
                        # Remove rows with invalid capacity data
                        self.data = self.data.dropna(subset=['Remaining capacity(Ah)'])
                        self.data = self.data[self.data['Remaining capacity(Ah)'] >= 0]
                        print(f"[INFO] Cleaned capacity data: {len(self.data)} rows remaining")
                    except Exception as e:
                        print(f"[ERROR] Failed to clean capacity data: {e}")
                
                # Clean cell voltage columns 
                try:
                    cell_voltage_cols = [col for col in self.data.columns[7:23] if col in self.data.columns]
                    for col in cell_voltage_cols:
                        if col in self.data.columns:
                            self.data[col] = self.data[col].astype(str)
                            self.data[col] = self.data[col].str.replace(',', '.')
                            self.data[col] = pd.to_numeric(self.data[col], errors='coerce')
                    
                    print(f"[INFO] Cleaned {len(cell_voltage_cols)} cell voltage columns")
                except Exception as e:
                    print(f"[ERROR] Failed to clean cell voltage data: {e}")
                
                # Reset index after filtering
                self.data = self.data.reset_index(drop=True)
                print(f"[INFO] Data cleaning complete: {len(self.data)} rows remaining")
            else:
                print("[WARNING] No required columns found for data cleaning")
                print(f"[INFO] Available columns: {list(self.data.columns)}")
                
                CTkMessagebox(
                    title="Column Error",
                    message=f"Required columns not found in the CSV file.\n\nExpected: {required_columns}\n\nFound: {list(self.data.columns)}\n\nPlease check the file format and column names.",
                    icon="warning"
                )

    def Alarms_Check(self, Start_index, Pre_End_index):

           # Load the config file
            with open("BAE_Alarms.json", "r") as file:
                config_data = json.load(file)

            # Extract the list of alarms from the config file
            # Liam CHeck
            alarms = config_data["Fail Alarms"]
            P1_alarms = config_data["Point 1 Alarms"]
            P2_alarms = config_data["Point 2 Alarms"]

            Failure_alarm_errors = []
            P1_alarm_errors = []
            P2_alarm_errors = []

            Test = True
            Test_P1 = False  # Start with False, set to True if ANY P1 alarm is found
            Test_P2 = False  # Start with False, set to True if ANY P2 alarm is found

            # Check for general failure alarms
            for index, row in self.data.iterrows():
                for alarm_key, alarm_value in alarms.items():
                    if alarm_value in str(row['Alarm']):
                        Failure_alarm_errors.append(f"Failure Alarm found in data at row {index + 1}: {alarm_value}")
                        Test = False

            if Failure_alarm_errors:
                self.errors.append(Failure_alarm_errors)
                self.alarms.append(Failure_alarm_errors)

            # Check for P1 alarms - need to find ALL alarms from ANY set around Start_index
            P1_alarm_sets = P1_alarms.values()
            Test_P1 = False
            for alarm_set in P1_alarm_sets:
                alarm_values = list(alarm_set.values())
                for offset in [-2, -1, 0, 1, 2]:
                    idx = Start_index + offset
                    if 0 <= idx < len(self.data):
                        row_alarm_str = str(self.data.iloc[idx]['Alarm'])
                        if all(alarm in row_alarm_str for alarm in alarm_values):
                            Test_P1 = True
                            break
                if Test_P1:
                    break

            if not Test_P1:
                P1_alarm_errors.append(
                    f"Missing trigger alarms for Point 1 found around row {Start_index + 1}. "
                    #f"Expected all alarms from any of these sets: {[[v for v in s.values()] for s in P1_alarms.values()]}"
                )
                self.errors.append(P1_alarm_errors)
                self.alarms.append(P1_alarm_errors)

            # Check for P2 alarms - need to find ALL alarms from ANY set around Pre_End_index
            P2_alarm_sets = P2_alarms.values()
            Test_P2 = False
            for alarm_set in P2_alarm_sets:
                alarm_values = list(alarm_set.values())
                for offset in [-2, -1, 0, 1, 2]:
                    idx = Pre_End_index + offset
                    if 0 <= idx < len(self.data):
                        row_alarm_str = str(self.data.iloc[idx]['Alarm'])
                        if all(alarm in row_alarm_str for alarm in alarm_values):
                            Test_P2 = True
                            break
                if Test_P2:
                    break

            if not Test_P2:
                P2_alarm_errors.append(
                    f"Missing trigger alarms for Point 2 found around row {Pre_End_index + 1}. "
                    #f"Expected all alarms from any of these sets: {[[v for v in s.values()] for s in P2_alarms.values()]}"
                )
                self.errors.append(P2_alarm_errors)
                self.alarms.append(P2_alarm_errors)

            # Overall test passes only if both P1 and P2 tests pass AND no failure alarms
            if not Test_P1 or not Test_P2:
                Test = False

            return Test

    # Display results in the text box
    def display_results(self):
        if self.data is None:
            self.result_text1.insert("end", "\nNo data loaded.\n")
            return
        
        # Load the config file
        with open("BAE_Config_Parameters.json", "r") as file:
            config_data = json.load(file)
            config_P1 = config_data["Point 1"]
            config_P2 = config_data["Point 2"]
            config_P3 = config_data["Point 3"]

        # Ensure the 'Battery voltage(V)' column contains string values
        self.data['Battery voltage(V)'] = self.data['Battery voltage(V)'].astype(str)
        self.data['Battery voltage(V)'] = self.data['Battery voltage(V)'].str.replace(',', '.').astype(float)

        # Ensure the 'Remaining capacity(Ah)' column contains string values
        self.data['Remaining capacity(Ah)'] = self.data['Remaining capacity(Ah)'].astype(str)
        self.data['Remaining capacity(Ah)'] = self.data['Remaining capacity(Ah)'].str.replace(',', '.').astype(float)

        # Find the row with the maximum battery voltage
        max_voltage_row = self.data['Battery voltage(V)'].idxmax()
        Start_index = max_voltage_row

        # Use the first row in the dataset as the End_index
        End_index = self.data.index[0]

        # Find the row with the minimum battery voltage
        min_voltage_row = self.data['Battery voltage(V)'].idxmin()
        Pre_End_index = min_voltage_row

        # Assume tests pass initially
        Concess_Test = False
        Final_Fail = False

        if self.Alarms_Check(Start_index, Pre_End_index):
            self.result_text1.insert("end", f"Alarms Test Passed.\n\n")
        else:
            self.result_text1.insert("end", f"Alarms Test Failed.\n")
            if self.alarms: 
                for alarm_group in self.alarms:
                    for alarm in alarm_group:
                        self.result_text1.insert("end", f"{alarm}\n")
                        self.Jsonalarms_errors.append(alarm)
                self.result_text1.insert("end", "\n") 
            else:
                self.result_text1.insert("end", "No specific alarms found.\n\n")


        for point_num in range(1, 4):
            point_key = f"Point {point_num}"

            # Extract parameters from the config
            BV = config_data[point_key][0]["Expected Battery voltage"]
            RC = config_data[point_key][0]["Expected Remaining Capacity"]
            Max_Cell_Diff = config_data[point_key][0]["Expected Max Cell Voltage Diff"]

            # Get the values from the Excel file
            # @ 100% Capacity (Max Voltage)
            if point_num == 1:
                row = self.data.loc[Start_index]
                Battery_Voltage = row['Battery voltage(V)']
                BV1 = Battery_Voltage
                Remaining_Capacity = float(row['Remaining capacity(Ah)'])
                RC1 = Remaining_Capacity
                Cell_Vol = [float(value.replace(',', '.')) if isinstance(value, str) else float(value) for value in row[7:23]]
                Max_Cell_Vol = max(Cell_Vol)
                Min_Cell_Vol = min(Cell_Vol)
                Cell_diff = Max_Cell_Vol - Min_Cell_Vol
                print(f"Point 1 Cell Voltages: {Cell_diff}")  
                CD1 = Cell_diff
            # @ 0% Capacity (Min Voltage)
            if point_num == 2:
                row = self.data.loc[Pre_End_index]
                Battery_Voltage = row['Battery voltage(V)']
                BV2 = Battery_Voltage
                Remaining_Capacity = float(row['Remaining capacity(Ah)'])
                RC2 = Remaining_Capacity
                Cell_Vol = [float(value.replace(',', '.')) if isinstance(value, str) else float(value) for value in row[7:23]]
                Max_Cell_Vol = max(Cell_Vol)
                Min_Cell_Vol = min(Cell_Vol)
                Cell_diff = Max_Cell_Vol - Min_Cell_Vol
                print(f"Point 2 Cell Voltages: {Cell_diff}")  
                CD2 = Cell_diff
            # @ 50% Capacity (Resting Voltage)
            if point_num == 3:
                row = self.data.loc[End_index]
                Battery_Voltage = row['Battery voltage(V)']
                BV3 = Battery_Voltage
                Remaining_Capacity = float(row['Remaining capacity(Ah)'])
                RC3 = Remaining_Capacity
                Cell_Vol = [float(value.replace(',', '.')) if isinstance(value, str) else float(value) for value in row[7:23]]
                Max_Cell_Vol = max(Cell_Vol)
                Min_Cell_Vol = min(Cell_Vol)
                Cell_diff = Max_Cell_Vol - Min_Cell_Vol
                print(f"Point 3 Cell Voltages: {Cell_diff}")  
                CD3 = Cell_diff

            # Collect failure messages per point
            point_failed = False
            point_concession = False
            fail_messages = []

            # Solid pass
            if (BV[0] <= Battery_Voltage <= BV[1] and
                RC[0] <= Remaining_Capacity <= RC[1] and
                Cell_diff <= Max_Cell_Diff[0]):
                Test = True
                self.result_text1.insert("end", f"{point_key} result -> Solid passed.\n")
            else:
                # Check Battery Voltage
                if not (BV[0] <= Battery_Voltage <= BV[1]):
                    self.errors.append(f'Battery Voltage at {point_key} not in range.\nReceived: {Battery_Voltage}, Range: [{BV[0]},{BV[1]}].')
                    self.Jsonalarms_errors.append(f'Battery Voltage at {point_key} not in range.\nReceived: {Battery_Voltage}, Range: [{BV[0]},{BV[1]}].')
                    fail_messages.append(f"Battery Voltage(V) at {point_key} NOT in range. Received: {Battery_Voltage}, Range: [{BV[0]} , {BV[1]}].\n")
                    Test = False
                    Final_Fail = True
                    point_failed = True

                # Check Remaining Capacity
                if not (RC[0] <= Remaining_Capacity <= RC[1]):
                    self.errors.append(f'Remaining Capacity at {point_key} not in range.\nReceived: {Remaining_Capacity}, Range: [{RC[0]},{RC[1]}]')
                    self.Jsonalarms_errors.append(f'Remaining Capacity at {point_key} not in range.\nReceived: {Remaining_Capacity}, Range: [{RC[0]},{RC[1]}]')
                    fail_messages.append(f"Remaining Capacity(Ah) at {point_key} NOT in range. Received: {Remaining_Capacity}, Range: [{RC[0]} , {RC[1]}].\n")
                    Test = False
                    Final_Fail = True
                    point_failed = True

                # Check Cell Voltage Difference
                if Cell_diff > Max_Cell_Diff[1]:
                    self.errors.append(f'Cell Diff at {point_key} not in range.\nReceived: {round(Cell_diff, 3)}, Range: [{Max_Cell_Diff[0]},{Max_Cell_Diff[1]}]')
                    self.Jsonalarms_errors.append(f'Cell Diff at {point_key} not in range.\nReceived: {round(Cell_diff, 3)}, Range: [{Max_Cell_Diff[0]},{Max_Cell_Diff[1]}]')
                    fail_messages.append(f"Cell Voltage Difference(V) at {point_key} NOT in range. Received: {round(Cell_diff, 3)}, Range: [{Max_Cell_Diff[0]} , {Max_Cell_Diff[1]}].\n")
                    Test = False
                    Final_Fail = True
                    point_failed = True

                # Concession pass 
                if (BV[0] <= Battery_Voltage <= BV[1] and
                    RC[0] <= Remaining_Capacity <= RC[1] and
                    Max_Cell_Diff[0] <= Cell_diff <= Max_Cell_Diff[1]):
                    self.result_text1.insert("end", f"{point_key} result -> Concession pass.\n")
                    self.concession_pass_points.append(f'Cell Diff at {point_key} in range.\nReceived: {round(Cell_diff, 3)}, Range: [{Max_Cell_Diff[0]},{Max_Cell_Diff[1]}]')
                    self.Jsonalarms_errors.append(f'Cell Diff at {point_key} in range.\nReceived: {round(Cell_diff, 3)}, Range: [{Max_Cell_Diff[0]},{Max_Cell_Diff[1]}]')
                    self.result_text1.insert("end", f"Cell Voltage Difference(V) {point_key} IN range. Received: {round(Cell_diff, 3)}, Range: [{Max_Cell_Diff[0]} , {Max_Cell_Diff[1]}].\n\n")
                    Test = True
                    Concess_Test = True

            # Print fail messages
            if point_failed and fail_messages:
                self.result_text1.insert("end", f"{point_key} result -> Fail test.\n")
                for msg in fail_messages:
                    self.result_text1.insert("end", msg)
                self.result_text1.insert("end", "\n")
     
            
        if (Test == True and self.Alarms_Check(Start_index, Pre_End_index) == True and Concess_Test == False and Final_Fail == False):

            CTkMessagebox(title="Pass", message="Test passed successfully!", icon="check")
            self.status_indicator.configure(text="PASS", text_color="green")
            self.result_text1.insert("end", f"\nFinal Result: Battery Passed ✅")
            self.result="Solid_Pass"
        elif(Test == True and self.Alarms_Check(Start_index, Pre_End_index) == True and Concess_Test == True and Final_Fail == False):

            CTkMessagebox(title="Concession Pass",message="Test passed with Concession",icon="warning") 
            self.status_indicator.configure(text="CONCESSION PASS", text_color="orange")
            self.result_text1.insert("end", f"\nFinal Result: Battery Passed with Concession ⚠️")
            self.result="Concession_Pass"
        else:
            CTkMessagebox(title="Fail", message="Test Failed", icon="cancel")
            self.status_indicator.configure(text="FAIL", text_color="red")
            error_message = "\n".join([str(error) for error in self.errors])
            self.result_text1.insert("end", f"\nFinal Result: Battery Aging Test Failed ❌")
            self.result="Fail"

        #print("\nResults:\n")
        
        #Test Results

        #Point 1#
        P1_Voltage_test = test("Point 1 Battery Voltage",str(((config_P1[0]["Expected Battery voltage"][0])+(config_P1[0]["Expected Battery voltage"][1]))/2), 
                    float(BV1), "Result", "Battery Voltage",str(config_P1[0]["Expected Battery voltage"]), "NA")
        
        P1_Capacity_test = test("Point 1 Remaining Capacity",str(((config_P1[0]["Expected Remaining Capacity"][0])+(config_P1[0]["Expected Remaining Capacity"][1]))/2), 
                    float(RC1), "Result", "Remaining Capacity",str(config_P1[0]["Expected Remaining Capacity"]), "NA")

        P1_Volatage_Diff = test("Point 1 Max Cell Voltage Diff",str(((config_P1[0]["Expected Max Cell Voltage Diff"][0])+(config_P1[0]["Expected Max Cell Voltage Diff"][1]))/2), 
                    float(CD1), "Result", "Max Cell Voltage Diff",str(config_P1[0]["Expected Max Cell Voltage Diff"]), "NA")
        
        self.tests.append(P1_Voltage_test.__dict__)
        self.tests.append(P1_Capacity_test.__dict__)
        self.tests.append(P1_Volatage_Diff.__dict__)
        #Point 1#

        #Point 2#
        P2_Voltage_test = test("Point 2 Battery Voltage",str(((config_P2[0]["Expected Battery voltage"][0])+(config_P2[0]["Expected Battery voltage"][1]))/2), 
                    float(BV2), "Result", "Battery Voltage",str(config_P2[0]["Expected Battery voltage"]), "NA")
        
        P2_Capacity_test = test("Point 2 Remaining Capacity",str(((config_P2[0]["Expected Remaining Capacity"][0])+(config_P2[0]["Expected Remaining Capacity"][1]))/2), 
                    float(RC2), "Result", "Remaining Capacity",str(config_P2[0]["Expected Remaining Capacity"]), "NA")

        P2_Volatage_Diff = test("Point 2 Max Cell Voltage Diff",str(((config_P2[0]["Expected Max Cell Voltage Diff"][0])+(config_P2[0]["Expected Max Cell Voltage Diff"][1]))/2), 
                    float(CD2), "Result", "Max Cell Voltage Diff",str(config_P2[0]["Expected Max Cell Voltage Diff"]), "NA")
        
        self.tests.append(P2_Voltage_test.__dict__)
        self.tests.append(P2_Capacity_test.__dict__)
        self.tests.append(P2_Volatage_Diff.__dict__)
        #Point 2#

        #Point 3#
        P3_Voltage_test = test("Point 3 Battery Voltage",str(((config_P3[0]["Expected Battery voltage"][0])+(config_P3[0]["Expected Battery voltage"][1]))/2), 
                    float(BV3), "Result", "Battery Voltage",str(config_P3[0]["Expected Battery voltage"]), "NA")
        
        P3_Capacity_test = test("Point 3 Remaining Capacity",str(((config_P3[0]["Expected Remaining Capacity"][0])+(config_P3[0]["Expected Remaining Capacity"][1]))/2), 
                    float(RC3), "Result", "Remaining Capacity",str(config_P3[0]["Expected Remaining Capacity"]), "NA")

        P3_Volatage_Diff = test("Point 3 Max Cell Voltage Diff",str(((config_P3[0]["Expected Max Cell Voltage Diff"][0])+(config_P3[0]["Expected Max Cell Voltage Diff"][1]))/2), 
                    float(CD3), "Result", "Max Cell Voltage Diff",str(config_P3[0]["Expected Max Cell Voltage Diff"]), "NA")
        
        self.tests.append(P3_Voltage_test.__dict__)
        self.tests.append(P3_Capacity_test.__dict__)
        self.tests.append(P3_Volatage_Diff.__dict__)    

        self.store_test_data()   
        #Point 3#

    # Display graph
    def display_graph1(self):
        if self.data is None:
            self.result_text1.insert("end", "\nNo data loaded.\n")
            return

        column_name = self.dropdown_menu1.get()

        # Check if the column exists in data
        if (column_name in self.data.columns or column_name == "Delta Cell Voltages"):
            if column_name == "Delta Cell Voltages":
                column_data = []
                for index, row in self.data.iterrows():
                    Cell_Vol = [float(value.replace(',', '.')) if isinstance(value, str) else float(value) for value in row[7:23]]
                    Max_Cell_Vol = max(Cell_Vol)
                    Min_Cell_Vol = min(Cell_Vol)
                    Cell_diff = Max_Cell_Vol - Min_Cell_Vol
                    column_data.append(Cell_diff)
            else:
                column_data = self.data[column_name].tolist()
        else:
            self.result_text1.insert("end", f"\nColumn '{column_name}' not found in the data.\n")
            return

        column_data_ed = [float(value.replace(',', '.')) if isinstance(value, str) else float(value) for value in column_data]
        reversed_column_data = list(reversed(column_data_ed))

        # Determine the min and max y values for scaling
        min_y = min(reversed_column_data)
        max_y = max(reversed_column_data)
        y_margin = (max_y - min_y) * 0.1  # 10% margin

        # Plot the data
        fig, ax = plt.subplots(figsize=(5, 4)) 
        line, = ax.plot(reversed_column_data, marker='o', linestyle='-', color='b', label=f'{column_name} Values')
        ax.set_ylim(min_y - y_margin, max_y + y_margin)
        ax.set_title(f'{column_name} vs Time', fontsize=14)
        ax.set_xlabel("Time", fontsize=12)
        ax.set_ylabel(f'{column_name}', fontsize=12)
        ax.legend()
        ax.grid(True)

        max_voltage_row = self.data['Battery voltage(V)'].idxmax()
        ax.axvline(x=len(self.data) - max_voltage_row - 1, color='r', linestyle='--')
        ax.legend()
        ax.axvspan(len(self.data) - max_voltage_row - 1, len(self.data), color='grey', alpha=0.5)


        # Function to display x and y values in a pop-up box on click
        def on_click(event):
            if event.button is MouseButton.LEFT:
                xdata, ydata = event.xdata, event.ydata
                if xdata is not None and ydata is not None:
                    messagebox.showinfo("Point Coordinates", f"x={xdata:.2f}, y={ydata:.2f}")

        fig.canvas.mpl_connect('button_press_event', on_click)

        # Display the plot in the Tkinter GUI
        for widget in self.graph_frame1.winfo_children():
            widget.destroy()

        canvas = FigureCanvasTkAgg(fig, master=self.graph_frame1)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        self.current_canvas = canvas  
        
    # Start Test
    def start_test(self):
        if not self.current_operator:
            CTkMessagebox(
                title="Access Denied",
                message="Please log in before starting a test",
                icon="warning"
            )
            self.show_login_dialog()
            pass
            return

        # Check if the DC_checkbox is ticked
        if not self.DC_checkbox.get():
            CTkMessagebox(
                title="Action Required",
                message="You can only commence with the test if the 'Data was Cleared' tick box is checked.",
                icon="warning"
            )
            return

        # Add validation before testing
        is_valid, message = self.validate_data_structure()
        if not is_valid:
            CTkMessagebox(
                title="Data Validation Error",
                message=f"Cannot run test: {message}",
                icon="cancel"
            )
            return

        # Set status to TESTING before running the test
        self.status_indicator.configure(text="Testing...", text_color="yellow")

        self.result_text1.insert("end", "Battery Criteria🔋\n")
        self.result_text1.insert("end", f"\nLoaded file: {self.device_id}\n\n")
        self.display_results()
        self.start_button.configure(state="disabled")
        self.reset_button.configure(state="normal")        
        #self.dropdown_menu1.configure(state="disabled")
        self.file_dropdown_menu.configure(state="disabled")
        # print(f'\nSelf.Alarms:\n{self.alarms}')
        # print(f'\nSelf.Errors:\n{self.errors}')
        # print(f'\nSelf.JsonAlarms:\n{self.Jsonalarms_errors}')
        # print(f'\nSelf.tests:\n{self.tests}')

    # Clear graph
    def clear_graph1(self):
        try:
            if self.current_canvas:
                plt.close(self.current_canvas.figure)
                self.current_canvas.get_tk_widget().destroy()
                self.current_canvas = None
            
            for widget in self.graph_frame1.winfo_children():
                widget.destroy()
        except Exception as e:
            print(f"Error during graph cleanup: {e}")

    # Reset GUI
    def reset_gui(self):
        self.result_text1.delete("1.0", "end")
        self.clear_graph1()
        self.start_button.configure(state="normal")
        self.reset_button.configure(state="normal")
        self.DC_checkbox.deselect() 
        self.status_indicator.configure(text="Ready", text_color="yellow")
        self.file_dropdown_menu.configure(state="normal")
        self.tests = []
        self.errors = []
        self.alarms = []
        self.Jsonalarms_errors = []
        self.concession_pass_points = []
        self.fail_point = []
        
        self.file_search_var.set("")
        
        # Refresh the dropdown menu 
        source = self.file_source_var.get()
        if source == "Local":
            self.populate_file_dropdown()
        else:  # SharePoint
            self.refresh_files()
        
        self.file_dropdown_menu.set("Select Excel File")
        self.start_button.configure(state="disabled")  
        
        if hasattr(self, 'loaded_filename'):
            delattr(self, 'loaded_filename')
        
        self.device_id = "NA"

    # Quit program
    def quit(self):
        try:
            self.quit_button.configure(state="disabled")
            self.start_button.configure(state="disabled")
            self.reset_button.configure(state="disabled")
            self.help_button.configure(state="disabled")
            self.display_button1.configure(state="disabled")
            self.clear_button1.configure(state="disabled")
            
            self.clear_graph1()
            plt.close('all')           

            self.update_idletasks()
            
            try:
                for after_id in self.tk.eval('after info').split():
                    self.after_cancel(after_id)
            except Exception:
                pass

            for widget in self.winfo_children():
                try:
                    widget.destroy()
                except Exception:
                    pass
            
            self.destroy()
            
        except Exception as e:
            print(f"Error during cleanup: {e}")
            # Force destroy as last resort
            try:
                self.destroy()
            except Exception:
                pass

# LOGIN and Out

    # Toggle Login
    def toggle_login(self):
        if self.current_operator:
            self.logout()
        else:
            self.show_login_dialog()

    def show_login_dialog(self):
        # Prevent multiple login windows
        if hasattr(self, 'login_window') and self.login_window is not None and self.login_window.winfo_exists():
            self.login_window.lift()  # Bring to front if already open
            return

        self.login_window = ctk.CTkToplevel()
        self.login_window.title("Operator Login")
        self.login_window.geometry("300x350")
        self.login_window.resizable(False, False)
        
        # Center the window
        x = self.winfo_x() + (self.winfo_width() // 2) - (300 // 2)
        y = self.winfo_y() + (self.winfo_height() // 2) - (350 // 2)
        self.login_window.geometry(f"400x450+{x}+{y}")
        
        # Main frame
        main_frame = ctk.CTkFrame(self.login_window)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Title
        title_label = ctk.CTkLabel(
            main_frame, 
            text="Operator Login", 
            font=ctk.CTkFont(size=20, weight="bold")
        )
        title_label.pack(pady=20)
        
        # Username
        name_label = ctk.CTkLabel(
            main_frame, 
            text="Username:",
            font=ctk.CTkFont(size=14)
        )
        name_label.pack(pady=(10, 5))
        
        name_entry = ctk.CTkEntry(
            main_frame,
            width=200,
            placeholder_text="Enter username"
        )
        name_entry.pack(pady =10)
        
        # Password
        password_label = ctk.CTkLabel(
            main_frame, 
            text="Password:",
            font=ctk.CTkFont(size=14)
        )
        password_label.pack(pady=(10, 5))
        
        password_entry = ctk.CTkEntry(
            main_frame,
            width=200,
            placeholder_text="Enter password",
            show="*"  # Mask password input
        )
        password_entry.pack(pady=10)

        def validate_login():
            # Checks User login, removes spaces and converts all to UpperCase
            time.sleep(0.1)
            entered_operator = ''.join(name_entry.get().split()).upper()
            time.sleep(0.1)
            entered_password = password_entry.get().strip()

            if entered_operator in self.configLogin['Users'] and self.configLogin['Users'][entered_operator] == entered_password:
                self.current_operator = entered_operator
                self.operator_label.configure(text=self.current_operator)
                self.lock_button.configure(fg_color="green")  # Change to green when logged in
                self.start_button.configure(state="normal")
                #self.inspection_checkbox.configure(state="normal")  # Enable the checkbox after login
                self.DC_checkbox.configure(state="normal")
                self.reset_button.configure(state="normal")
                self.status_indicator.configure(text="Ready", text_color="yellow")
                # Enable all panels after login
                self.set_panels_state("normal")
                self.login_window.destroy()
            else:
                messagebox.showwarning("Log in Error", "Log In Error, please try again")
                name_entry.delete(0, 'end')
                password_entry.delete(0, 'end')

        # Sign In button 
        sign_in_button = ctk.CTkButton(
            main_frame,
            text="Sign In",
            command=validate_login,
            width=200,
            height=40,
            font=ctk.CTkFont(size=14, weight="bold")
        )
        sign_in_button.pack(pady=(30, 10))
        
        # Cancel button 
        cancel_button = ctk.CTkButton(
            main_frame,
            text="Cancel",
            command=self.login_window.destroy,
            width=200,
            height=40,
            fg_color="transparent",
            border_width=2,
            font=ctk.CTkFont(size=14)
        )
        cancel_button.pack(pady=(0, 10))
        
        def on_closing():
            self.login_window.destroy()
            self.login_window = None  # Reset reference

        self.login_window.protocol("WM_DELETE_WINDOW", on_closing)
        
        self.login_window.transient(self)

        name_entry.focus_set()
        
        name_entry.bind("<Return>", lambda event: validate_login())

    # Logout
    def logout(self):
        self.current_operator = None
        # self.serial_entry.delete(0, "end")
        # self.serial_entry2.delete(0, "end")
        self.operator_label.configure(text="Not logged in")
        self.lock_button.configure(fg_color="red")  # Change to red when logged out
        self.start_button.configure(state="disabled")  # Disable start button
        self.reset_button.configure(state="normal")  # Disable reset button
        self.status_indicator.configure(text="Not logged in", text_color="yellow")
        #self.inspection_checkbox.deselect()
        #self.inspection_checkbox.configure(state="disabled")
        self.DC_checkbox.deselect()
        self.DC_checkbox.configure(state="disabled")
        # Disable all panels after logout
        self.set_panels_state("disabled")

    # Validate Credentials
    def validate_credentials(self, username, pin):
        users = self.configLogin.get('Users', {})
        return username in users and users[username] == pin

    def reload_sharepoint_config(self):
        try:
            with open('BAE_SharePoint_Config.json', 'r') as f:
                sharepoint_config_data = json.load(f)
                self.sharepoint_config = sharepoint_config_data["SharePoint_Config"]
                print(f"[INFO] Reloaded SharePoint config - new folder: {self.sharepoint_config['folder_url']}")
                return True
        except FileNotFoundError:
            print("[WARNING] SharePoint config file not found, using existing values")
            return False
        except Exception as e:
            print(f"[ERROR] Failed to reload SharePoint config: {e}")
            return False

    # SharePoint Integration Methods
    def get_sharepoint_access_token(self):
        """Get access token for Microsoft Graph API."""
        try:
            graph_config = self.sharepoint_config.get('graph_api', {})
            if not all(key in graph_config for key in ['client_id', 'client_secret', 'tenant_id']):
                print("[ERROR] Graph API configuration incomplete")
                return None
            
            app = ConfidentialClientApplication(
                client_id=graph_config['client_id'],
                client_credential=graph_config['client_secret'],
                authority=f"https://login.microsoftonline.com/{graph_config['tenant_id']}"
            )
            
            scopes = ['https://graph.microsoft.com/.default']
            result = app.acquire_token_for_client(scopes=scopes)
            
            if 'access_token' not in result:
                error_desc = result.get('error_description', 'Unknown error')
                print(f"[ERROR] Authentication failed: {error_desc}")
                return None
            
            print("[SUCCESS] Graph API authentication successful")
            return result['access_token']
            
        except Exception as e:
            print(f"[ERROR] Failed to get access token: {e}")
            return None

    def get_sharepoint_files(self):
        """Get list of Excel files from SharePoint folder using Graph API."""
        if not SHAREPOINT_AVAILABLE:
            CTkMessagebox(
                title="SharePoint Error",
                message="SharePoint libraries not installed. Please install: pip install msal requests",
                icon="cancel"
            )
            return []

        try:
            print(f"[INFO] Getting SharePoint files using Graph API...")
            
            # Get access token
            access_token = self.get_sharepoint_access_token()
            if not access_token:
                return []
            
            headers = {'Authorization': f'Bearer {access_token}'}
            graph_config = self.sharepoint_config.get('graph_api', {})
            qa_drive_id = graph_config.get('qa_drive_id')
            
            if not qa_drive_id:
                print("[ERROR] QA Drive ID not configured")
                return []
            
            # Convert folder URL to Graph API path
            folder_path = self.sharepoint_config['folder_url']
            # Remove site prefix if present
            if folder_path.startswith('/sites/Assurance-Quality'):
                folder_path = folder_path.replace('/sites/Assurance-Quality', '')
            
            # Get files from folder
            folder_url = f"https://graph.microsoft.com/v1.0/drives/{qa_drive_id}/root:{folder_path}:/children"
            print(f"[DEBUG] Folder URL: {folder_url}")
            
            response = requests.get(folder_url, headers=headers)
            
            if response.status_code != 200:
                print(f"[ERROR] Failed to get folder contents: {response.status_code}")
                if response.text:
                    print(f"[DEBUG] Response: {response.text[:500]}")
                return []
            
            data = response.json()
            excel_files = []
            
            for item in data.get('value', []):
                if item.get('file'):  # It's a file, not a folder
                    file_name = item.get('name', '')
                    if file_name.endswith(('.xlsx', '.xls', '.csv')):
                        excel_files.append(file_name)
            
            print(f"[SUCCESS] Found {len(excel_files)} Excel files on SharePoint")
            return excel_files
            
        except Exception as e:
            print(f"[ERROR] SharePoint connection error: {e}")
            CTkMessagebox(
                title="Connection Error",
                message=f"SharePoint connection error: {str(e)}",
                icon="cancel"
            )
            return []

    def load_sharepoint_file(self, file_name):
        """Load an Excel file from SharePoint using Graph API."""
        if not SHAREPOINT_AVAILABLE:
            return None

        try:
            print(f"[INFO] Loading file from SharePoint using Graph API: {file_name}")
            
            # Get access token
            access_token = self.get_sharepoint_access_token()
            if not access_token:
                return None
            
            headers = {'Authorization': f'Bearer {access_token}'}
            graph_config = self.sharepoint_config.get('graph_api', {})
            qa_drive_id = graph_config.get('qa_drive_id')
            
            if not qa_drive_id:
                print("[ERROR] QA Drive ID not configured")
                return None
            
            # Convert folder URL to Graph API path
            folder_path = self.sharepoint_config['folder_url']
            if folder_path.startswith('/sites/Assurance-Quality'):
                folder_path = folder_path.replace('/sites/Assurance-Quality', '')
            
            # Construct full file path
            file_path = f"{folder_path}/{file_name}"
            file_url = f"https://graph.microsoft.com/v1.0/drives/{qa_drive_id}/root:{file_path}:/content"
            print(f"[DEBUG] File URL: {file_url}")
            
            # Download file
            response = requests.get(file_url, headers=headers)
            
            if response.status_code != 200:
                print(f"[ERROR] Failed to download file: {response.status_code}")
                if response.text:
                    print(f"[DEBUG] Response: {response.text[:200]}")
                CTkMessagebox(
                    title="File Error",
                    message=f"Failed to download file from SharePoint: {response.status_code}",
                    icon="cancel"
                )
                return None
            
            print(f"[SUCCESS] File downloaded ({len(response.content)} bytes)")
            
            # Create BytesIO object from response
            bytes_file_obj = io.BytesIO(response.content)
            
            # Load the file based on its extension
            try:
                if file_name.endswith(".csv"):
                    # Try multiple CSV reading strategies like in local file loading
                    csv_strategies = [
                        {'sep': ';', 'encoding': 'utf-8'},
                        {'sep': ';', 'encoding': 'utf-8-sig'},
                        {'sep': ',', 'encoding': 'utf-8'},
                        {'sep': ',', 'encoding': 'utf-8-sig'}
                    ]
                    
                    df = None
                    for strategy in csv_strategies:
                        try:
                            bytes_file_obj.seek(0)  # Reset stream position
                            df = pd.read_csv(bytes_file_obj, **strategy)
                            if len(df.columns) > 1:
                                print(f"[SUCCESS] CSV loaded with strategy: {strategy}")
                                break
                        except Exception:
                            continue
                    
                    if df is None:
                        raise Exception("All CSV loading strategies failed")
                else:  # Excel files
                    df = pd.read_excel(bytes_file_obj)
                
                print(f"[SUCCESS] File loaded into DataFrame with {len(df)} rows and {len(df.columns)} columns")
                return df
                
            except Exception as file_error:
                print(f"[ERROR] Failed to parse file: {file_error}")
                CTkMessagebox(
                    title="File Error",
                    message=f"Failed to parse file: {str(file_error)}",
                    icon="cancel"
                )
                return None
                
        except Exception as e:
            print(f"[ERROR] Error loading SharePoint file: {e}")
            CTkMessagebox(
                title="Error",
                message=f"Error loading SharePoint file: {str(e)}",
                               icon="cancel"
            )
            return None

    def upload_file_to_sharepoint(self, local_file_path, sharepoint_folder_path, file_name):
        """Upload a file to SharePoint using Graph API."""
        if not SHAREPOINT_AVAILABLE:
            print("[WARNING] SharePoint not available, skipping upload")
            return False

        try:
            print(f"[INFO] Uploading file to SharePoint using Graph API: {file_name}")
            
            # Get access token
            access_token = self.get_sharepoint_access_token()
            if not access_token:
                return False
            
            headers = {'Authorization': f'Bearer {access_token}'}
            graph_config = self.sharepoint_config.get('graph_api', {})
            qa_drive_id = graph_config.get('qa_drive_id')
            
            if not qa_drive_id:
                print("[ERROR] QA Drive ID not configured")
                return False
            
            # Convert SharePoint folder path to Graph API path
            if sharepoint_folder_path.startswith('/sites/Assurance-Quality'):
                sharepoint_folder_path = sharepoint_folder_path.replace('/sites/Assurance-Quality', '')
            
            # Read the local file
            with open(local_file_path, 'rb') as file_content:
                file_data = file_content.read()
            
            # Upload file
            upload_url = f"https://graph.microsoft.com/v1.0/drives/{qa_drive_id}/root:{sharepoint_folder_path}/{file_name}:/content"
            print(f"[DEBUG] Upload URL: {upload_url}")
            
            upload_response = requests.put(upload_url, headers=headers, data=file_data)
            
            if upload_response.status_code in [200, 201]:
                print(f"[SUCCESS] File uploaded to SharePoint: {file_name}")
                return True
            else:
                print(f"[ERROR] Failed to upload file: {upload_response.status_code}")
                if upload_response.text:
                    print(f"[DEBUG] Response: {upload_response.text[:200]}")
                return False
                
        except Exception as e:
            print(f"[ERROR] Error uploading file to SharePoint: {e}")
            return False

    def delete_sharepoint_file(self, file_name):
        """Delete a file from SharePoint using Graph API."""
        if not SHAREPOINT_AVAILABLE:
            print("[WARNING] SharePoint not available, cannot delete file")
            return False

        try:
            print(f"[INFO] Deleting file from SharePoint: {file_name}")
            
            # Get access token
            access_token = self.get_sharepoint_access_token()
            if not access_token:
                return False
            
            headers = {'Authorization': f'Bearer {access_token}'}
            graph_config = self.sharepoint_config.get('graph_api', {})
            qa_drive_id = graph_config.get('qa_drive_id')
            
            if not qa_drive_id:
                print("[ERROR] QA Drive ID not configured")
                return False
            
            # Convert folder URL to Graph API path
            folder_path = self.sharepoint_config['folder_url']
            if folder_path.startswith('/sites/Assurance-Quality'):
                folder_path = folder_path.replace('/sites/Assurance-Quality', '')
            
            # Construct full file path
            file_path = f"{folder_path}/{file_name}"
            delete_url = f"https://graph.microsoft.com/v1.0/drives/{qa_drive_id}/root:{file_path}"
            print(f"[DEBUG] Delete URL: {delete_url}")
            
            # Delete file
            response = requests.delete(delete_url, headers=headers)
            
            if response.status_code == 204:  # No Content - successful deletion
                print(f"[SUCCESS] File deleted from SharePoint: {file_name}")
                return True
            elif response.status_code == 404:  # Not Found - file doesn't exist
                print(f"[WARNING] File not found on SharePoint (may already be deleted): {file_name}")
                return True  # Consider this successful since the goal is achieved
            else:
                print(f"[ERROR] Failed to delete file from SharePoint: {response.status_code}")
                if response.text:
                    print(f"[DEBUG] Response: {response.text[:200]}")
                return False
                
        except Exception as e:
            print(f"[ERROR] Error deleting SharePoint file: {e}")
            return False

    def upload_test_results_to_sharepoint(self, result_type, json_file_path, excel_file_path, device_id):
        """Upload test results to SharePoint based on test outcome."""
        if not SHAREPOINT_AVAILABLE:
            print("[WARNING] SharePoint not available, skipping result upload")
            return

        try:
            # Reload config to get upload paths
            self.reload_sharepoint_config()
            
            # Check if uploads are enabled
            if not self.sharepoint_config.get('enable_uploads', False):
                print("[INFO] SharePoint uploads disabled in configuration")
                return
            
            if 'upload_paths' not in self.sharepoint_config:
                print("[WARNING] Upload paths not configured in SharePoint config")
                return

            upload_config = self.sharepoint_config['upload_paths']
            
            # Determine folder based on result type
            if result_type == "Solid_Pass":
                folder_suffix = upload_config['pass_folder']
            elif result_type == "Concession_Pass":
                folder_suffix = upload_config['concession_pass_folder']
            elif result_type == "Fail":
                folder_suffix = upload_config['fail_folder']
            else:
                print(f"[WARNING] Unknown result type: {result_type}")
                return

            print(f"[INFO] Starting SharePoint upload for {result_type} results...")

            # Upload JSON file
            if json_file_path and os.path.exists(json_file_path):
                json_folder = f"{upload_config['json_base']}/{folder_suffix}"
                json_filename = os.path.basename(json_file_path)
                
                success = self.upload_file_to_sharepoint(json_file_path, json_folder, json_filename)
                if success:
                    print(f"[SUCCESS] JSON file uploaded to SharePoint: {json_folder}/{json_filename}")
                else:
                    print(f"[ERROR] Failed to upload JSON file to SharePoint")
            else:
                print(f"[WARNING] JSON file not found or path invalid: {json_file_path}")

            # Upload Excel file
            if excel_file_path and os.path.exists(excel_file_path):
                excel_folder = f"{upload_config['excel_base']}/{folder_suffix}"
                excel_filename = os.path.basename(excel_file_path)
                
                success = self.upload_file_to_sharepoint(excel_file_path, excel_folder, excel_filename)
                if success:
                    print(f"[SUCCESS] Excel file uploaded to SharePoint: {excel_folder}/{excel_filename}")
                else:
                    print(f"[ERROR] Failed to upload Excel file to SharePoint")
            else:
                print(f"[WARNING] Excel file not found or path invalid: {excel_file_path}")

            print(f"[INFO] SharePoint upload process completed for {result_type}")

        except Exception as e:
            print(f"[ERROR] Error uploading test results to SharePoint: {e}")
            CTkMessagebox(
                title="Upload Error",
                message=f"Failed to upload results to SharePoint: {str(e)}",
                icon="warning"
            )
    def validate_data_structure(self):
        """Validate that the data has the expected structure for testing."""
        if self.data is None:
            return False, "No data loaded"
        
        required_columns = ['Battery voltage(V)', 'Remaining capacity(Ah)', 'Alarm']
        missing_columns = [col for col in required_columns if col not in self.data.columns]
        if missing_columns:
            return False, f"Missing required columns: {missing_columns}"
        
        # Check for reasonable data ranges - handle both string and numeric data
        try:
            voltage_column = self.data['Battery voltage(V)']
            
            # Check if data is already numeric
            if pd.api.types.is_numeric_dtype(voltage_column):
                voltage_data = voltage_column
            else:
                # Convert string data with comma decimal separators to numeric
                voltage_data = pd.to_numeric(voltage_column.astype(str).str.replace(',', '.'), errors='coerce')
            
            if voltage_data.isna().all():
                return False, "Invalid voltage data - all values are NaN"
            
            # Check if we have enough valid data points
            valid_voltage_count = voltage_data.notna().sum()
            if valid_voltage_count < 10:
                return False, f"Insufficient valid voltage data points: {valid_voltage_count} (need at least 10)"
        
        except Exception as e:
            return False, f"Error processing voltage data: {str(e)}"
        
        # Check if we have enough data points
        if len(self.data) < 10:
            return False, "Insufficient data points for analysis"
        
        # Check if max and min voltages are reasonable
        try:
            max_voltage = voltage_data.max()
            min_voltage = voltage_data.min()
            voltage_range = max_voltage - min_voltage
            
            if pd.isna(max_voltage) or pd.isna(min_voltage):
                return False, "Unable to determine voltage range - invalid voltage values"
            
            if voltage_range < 5:  # Expect at least 5V difference between max and min
                return False, f"Voltage range too small: {voltage_range:.2f}V (expected at least 5V)"
                
        except Exception as e:
            return False, f"Error calculating voltage range: {str(e)}"
        
        return True, "Data structure valid"

# Main function to run app
def main():
    try:
        import sys
        if hasattr(sys, 'setrecursionlimit'):
            sys.setrecursionlimit(2000)
        app = BAETestApp()
        if hasattr(app, 'attributes'):
            try:
                app.attributes('-dpi', 96)
            except Exception:
                pass
        app.mainloop()
    except KeyboardInterrupt:
        if 'app' in locals():
            app.quit()
    except Exception as e:
        print(f"An error occurred: {e}")
        import traceback
        traceback.print_exc()
    finally:
        try:
            import gc
            gc.collect()
        except:
            pass

if __name__ == "__main__":
    main()
