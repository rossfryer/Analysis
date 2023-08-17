import os
import tkinter as tk
from tkinter import ttk, filedialog
import tkinter.messagebox as messagebox
from pandastable import Table
from tkinter import scrolledtext
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
import pandas as pd
import pymongo
import seaborn as sns


class TabbedInterface(tk.Tk):
    def __init__(self):
        super().__init__()

        self.text_area1 = None
        self.canvas1 = None
        self.text_area2 = None
        self.canvas2 = None
        self.nested_notebookDA = None
        self.tab1NestedDA = None
        self.tab2NestedDA = None
        self.tab3NestedDA = None
        self.text_area3 = None
        self.canvasGraph = None
        self.canvasCorr = None
        self.pt = None
        self.pt = None
        self.cleaned_Params = None
        self.cleaned_Ants = None
        self.pt = None
        self.file_entry1 = None
        self.file_entry2 = None
        self.merged = pd.DataFrame
        self.canvas = None
        self.canvasMMM = None
        self.title("DAB Analysis")
        self.geometry("1000x800")
        menu = tk.Menu(self)
        self.config(menu=menu)
        self.tab1_action_completed1 = False
        self.tab1_action_completed2 = False


        fileMenu = tk.Menu(menu)
        fileMenu.add_command(label="Exit", command=self.on_exit)
        menu.add_cascade(label="File", menu=fileMenu)

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self.notebook.bind("<<NotebookTabChanged>>", self.tab_changed)

        # Load and Clean Data Tab
        self.tab1 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab1, text="Load and Clean Data      ")
        self.load_and_clean(self.tab1)

        # Extract Data Tab
        self.tab2 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab2, text="Extract Data              ")
        self.extract_data(self.tab2)

        # Data Analysis Tab
        self.tab3 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab3, text="Data Analysis              ")
        self.data_analysis(self.tab3)

    def on_exit(self):
        result = tk.messagebox.askyesno("Exit", "Are you sure you want to exit?")
        if result:
            self.destroy()

    def load_and_clean(self,tab):
        text_box_frameR = ttk.Frame(tab, relief="raised", padding=(10, 10))
        text_box_frameR.grid(row=0, column=0, columnspan=4, sticky="w")

        text_boxR = tk.Text(tab, bd=1, relief=tk.SOLID, height=7)
        text_boxR.grid(row=0, column=0, columnspan=4,  padx=20,sticky="w")
        text_boxR.tag_configure('heading', font=('Arial', 12, 'bold'))
        text_boxR.tag_configure('main', font=('Arial', 10))

        # Insert the static text into the Text widget
        title = "Load and clean the initial dataset \n"
        mainText = "Select the CSV file to load, Clean the data, and then Save. \nThe data will then be converted to JSON format and added as a Collection to the Database."
        text_boxR.insert(tk.END, title, 'heading')
        text_boxR.insert(tk.END, mainText, 'main')

        text_box_frameL = ttk.Frame(tab, relief="raised", padding=(10, 10))
        text_box_frameL.grid(row=0, column=6, columnspan=2, pady=20, sticky="ew")

        text_boxL = tk.Text(tab, bd=1, relief=tk.SOLID, width = 30, height=7)
        text_boxL.grid(row=0, column=6, columnspan=2, pady=20 )
        text_boxL.tag_configure('heading', font=('Arial', 12, 'bold'))
        text_boxL.tag_configure('main', font=('Arial', 10))

        # Insert the static text into the Text widget
        title = "Save JSON to database\n"
        mainText = "When you have loaded, viewed and corrected your CSV files, save to Database."
        text_boxL.insert(tk.END, title, 'heading')
        text_boxL.insert(tk.END, mainText, 'main')

        # Buttons to Save the CSV files
        text_box_frameS = ttk.Frame(tab, relief="raised",  padding=(10, 10))
        text_box_frameS.grid(row=2, column=6, columnspan=1, pady=20)
        file_label = ttk.Label(tab, text="Save to database:")
        file_label.grid(row=2, column=6, sticky="w")
        save_button1 = ttk.Button(tab, text="Save", command=lambda: self.save_file(self.file_entry1))
        save_button1.grid(row=2, column=7, sticky='e')
        file_label = ttk.Label(tab, text="Save to database:")
        file_label.grid(row=3, column=6, sticky="w")
        save_button2 = ttk.Button(tab, text="Save", command=lambda: self.save_file(self.file_entry2))
        save_button2.grid(row=3, column=7, sticky="e")


        # First file selection picker
        file_label = ttk.Label(tab, text="Selected File 1:")
        file_label.grid(row=1, column=0,  padx=20, sticky="w")
        self.file_entry1 = ttk.Entry(tab, width=60, state="readonly")
        self.file_entry1.grid(row=2, column=0, padx=20, sticky="w")
        pick_button1 = ttk.Button(tab, text="Load File", command=lambda: self.select_file(self.file_entry1))
        pick_button1.grid(row=2, column=1, sticky="w")
        view_button1 = ttk.Button(tab, text="View File", command=lambda: self.display_recordsAnt_with_validation(self.file_entry1.get()))
        view_button1.grid(row=2, column=2,   sticky="w")
        clean_button1 = ttk.Button(tab, text="Clean File", command=lambda: self.clean_recordsAnt_with_validation(self.file_entry1.get()))
        clean_button1.grid(row=2, column=3,   sticky="w")

        # Second file selection picker
        file_label = ttk.Label(tab, text="Selected File 2:")
        file_label.grid(row=3, column=0, padx=20, sticky="w")
        self.file_entry2 = ttk.Entry(tab, width=60, state="readonly")
        self.file_entry2.grid(row=4, column=0,  padx=20, sticky="w")
        pick_button2 = ttk.Button(tab, text="Load File", command=lambda: self.select_file(self.file_entry2))
        pick_button2.grid(row=4, column=1,sticky="w")
        view_button2 = ttk.Button(tab, text="View File", command=lambda: self.display_recordsParam(self.file_entry2.get()))
        view_button2.grid(row=4, column=2,sticky="w")
        clean_button2 = ttk.Button(tab, text="Clean File", command=lambda: self.clean_recordsParam(self.file_entry2.get()))
        clean_button2.grid(row=4, column=3,sticky="w")

        # Text area to display feedback to user
        self.text_area1 = scrolledtext.ScrolledText(tab, height=10)
        self.text_area1.grid(row=5, column=0, columnspan=4, padx=10, pady=10)

        # Add some text to the text area
        self.text_area1.insert(tk.END, "This is a sample text.")

        # Create a Scrollbar
        scrollbar = tk.Scrollbar(tab, command=self.text_area1.yview)
        scrollbar.grid(row=5, column=5, sticky="w")

        # Attach the Scrollbar to the Text widget
        self.text_area1.configure(yscrollcommand=scrollbar.set)

        # Canvas to display the Pandastable of the data
        self.canvas1 = tk.Frame(tab, width=900, height=300, borderwidth=0)
        self.canvas1.grid(row=8, column=0, columnspan=8, padx=20, pady=20, sticky="nsew")

    # Tabbed area to Extract data and save Merged file
    def extract_data(self, tab):
        # Display text area at top of Tabbed area
        text_box_frameR = ttk.Frame(tab, relief="raised", padding=(10, 10))
        text_box_frameR.grid(row=0, column=0, columnspan=4, sticky="w")

        text_boxR = tk.Text(tab, bd=1, relief=tk.SOLID, height=7)
        text_boxR.grid(row=0, column=0, columnspan=4, padx=20, sticky="w")
        text_boxR.tag_configure('heading', font=('Arial', 12, 'bold'))
        text_boxR.tag_configure('main', font=('Arial', 10))

        # Insert the static text into the Text widget
        title = "Extract relevant data from dataset\n"
        mainText = "Data will be merged into a selection based on the following criteria: \n1.	Outputs should not include any data from DAB Radio stations that have the following ‘NGR’ : \nNZ02553847, SE213515, NT05399374 and NT252675908\n2.	The ‘EID’ column contains information of the DAB multiplex block E.g C19A. Extract this out into a new column, one for each of the following DAB multiplexes:\na.	all DAB multiplexes, that are , C18A, C18F, C188\nb.	join each category, C18A, C18F, C188 to the ‘ NGR’ that signifies the DAB stations location to the following: ‘Site’, ‘Site Height, In-Use Ae Ht, In-Use ERP Total \nc.	Please note that: In-Use Ae Ht, In-Use ERP Total  will need the following new header after extraction: Aerial height(m), Power(kW) respectively."
        text_boxR.insert(tk.END, title, 'heading')
        text_boxR.insert(tk.END, mainText, 'main')

        text_box_frameL = ttk.Frame(tab, relief="raised", padding=(10, 10))
        text_box_frameL.grid(row=0, column=6, columnspan=2, pady=20, sticky="ew")

        text_boxL = tk.Text(tab, bd=1, relief=tk.SOLID, width=30, height=7)
        text_boxL.grid(row=0, column=6, columnspan=2, pady=20)
        text_boxL.tag_configure('heading', font=('Arial', 12, 'bold'))
        text_boxL.tag_configure('main', font=('Arial', 10))

        # Insert the static text into the Text widget
        title = "Extract data\n"
        mainText = "When you are ready, press Extract below to prepare your new dataset. Save when you are ready."
        text_boxL.insert(tk.END, title, 'heading')
        text_boxL.insert(tk.END, mainText, 'main')

        # Text area to contain buttons to extract and save data
        text_box_frameS = ttk.Frame(tab, relief="raised", padding=(10, 10))
        text_box_frameS.grid(row=2, column=6, columnspan=1, padx=20, pady=20)
        file_label = ttk.Label(tab, text="Extract data:")
        file_label.grid(row=2, column=6, sticky="w")
        extract_button1 = ttk.Button(tab, text="Extract", command=lambda: self.extract_data_frames())
        extract_button1.grid(row=2, column=7, sticky='e')
        file_label = ttk.Label(tab, text="Save extracted data:")
        file_label.grid(row=3, column=6, sticky="w")
        save_button2 = ttk.Button(tab, text="Save", command=lambda: self.save_extracted_data())
        save_button2.grid(row=3, column=7, sticky="e")

        # Test area to supply feedback to user
        self.text_area2 = scrolledtext.ScrolledText(tab, height=10)
        self.text_area2.grid(row=4, column=0, columnspan=4, padx=10, pady=10)

        # Add some text to the text area
        self.text_area2.insert(tk.END, "This is a sample text.")

        # Create a Scrollbar
        scrollbar = tk.Scrollbar(tab, command=self.text_area2.yview)
        scrollbar.grid(row=4, column=5, sticky="w")

        # Attach the Scrollbar to the Text widget
        self.text_area2.configure(yscrollcommand=scrollbar.set)

        self.canvas2 = tk.Canvas(tab, width=900, height=300, borderwidth=0)
        self.canvas2.grid(row=5, column=0, columnspan=8, padx=20, pady=20, sticky="nsew")

    # Tabbed area to display the results of the required analysis
    def data_analysis(self, tab):
        text_box_frameR = ttk.Frame(tab, relief="raised", padding=(10, 10))
        text_box_frameR.grid(row=0, column=0, columnspan=4, sticky="w")

        text_boxR = tk.Text(tab, bd=1, relief=tk.SOLID, height=7)
        text_boxR.grid(row=0, column=0, columnspan=4, padx=20, pady=20, sticky="w")
        text_boxR.tag_configure('heading', font=('Arial', 12, 'bold'))
        text_boxR.tag_configure('main', font=('Arial', 10))

        # Insert the static text into the Text widget
        title = "Analysis of Data\n"
        mainText = "Choose the analysis that you require.\nMean, median and mode for the extracted EIDs.\nGraph of extracted data.\nCorrelation of extracted data.\n"
        text_boxR.insert(tk.END, title, 'heading')
        text_boxR.insert(tk.END, mainText, 'main')

        # Nested Notebook for Data Analysis - three tabs
        self.nested_notebookDA = ttk.Notebook(tab)
        self.nested_notebookDA.grid(row=1, column=0, columnspan=4, padx=20, sticky="w")

        # Tab 1 - Mean Median and Mode
        self.tab1NestedDA = ttk.Frame(self.nested_notebookDA)
        self.nested_notebookDA.add(self.tab1NestedDA, text="Mean, Median and Mode                    ")
        self.data_analysis_mean(self.tab1NestedDA)

        # Tab 2 - Graphs
        self.tab2NestedDA = ttk.Frame(self.nested_notebookDA)
        self.nested_notebookDA.add(self.tab2NestedDA, text="Graphical Analysis                        ")
        self.data_analysis_graph(self.tab2NestedDA)

        # Tab 3 - Correlation
        self.tab3NestedDA = ttk.Frame(self.nested_notebookDA)
        self.nested_notebookDA.add(self.tab3NestedDA, text="Correlation                               ")
        self.data_analysis_correlation(self.tab3NestedDA)

    # Contents of TAB 1
    def data_analysis_mean(self, tab):
        label = ttk.Label(tab, text="\nEIDs extracted:\nC118, C11, C18A\n")
        label.grid(row=3, column=0, padx=20, sticky="w")

        # Button to analyse extracted data for selected criteria
        analyse = ttk.Button(tab, text="Analyse", command=lambda: self.display_records_analysis(tab))
        analyse.grid(row=6, column=0, padx=20, sticky="w")

        self.text_area3 = scrolledtext.ScrolledText(tab, height=10)
        self.text_area3.grid(row=7, column=0, columnspan=4, padx=10, pady=10)

        # Add some text to the text area
        self.text_area3.insert(tk.END, "")

        # Create a Scrollbar
        scrollbar = tk.Scrollbar(tab, command=self.text_area3.yview)
        scrollbar.grid(row=7, column=5, sticky="w")

        # Attach the Scrollbar to the Text widget
        self.text_area3.configure(yscrollcommand=scrollbar.set)

        self.canvasMMM = tk.Canvas(tab, width=900, height=300, borderwidth=0)
        self.canvasMMM.grid(row=8, column=0, columnspan=8, padx=20, pady=20, sticky="nsew")

    # Carries out required Mean, Median and Mode analysis and Merged data and displays results
    def display_records_analysis(self,tab):
        # Calculation for Site Height > 75
        df = self.merged
        df_siteHeight = df[df['Site Height'] > 75]
        mean_value_siteHeight = df_siteHeight['Power(kW)'].mean().round(3)
        median_value_siteHeight = df_siteHeight['Power(kW)'].median()
        mode_value_siteHeight = df_siteHeight['Power(kW)'].mode().iloc[0]

        # Calculation for Date > 2001
        df_date = df[df['Date'] > '2001-01-01']
        mean_value_date = df_date['Power(kW)'].mean().round(3)
        median_value_date = df_date['Power(kW)'].median()
        mode_value_date = df_date['Power(kW)'].mode().iloc[0]

        # Display of results
        siteHeightResults = f'For Site Height greater than 75, mean is {mean_value_siteHeight}, median is {median_value_siteHeight} and the mode is {mode_value_siteHeight}'
        dateResults = f'For Date from 2001, mean is {mean_value_date}, median is {median_value_date} and the mode is {mode_value_date}'
        self.text_area3.insert(tk.END, f'{siteHeightResults}\n\n{dateResults}')
        # Create the treeview

        tree = ttk.Treeview(self.canvasMMM)

        # Define columns
        tree["columns"] = ("Condition", "Mean", "Median", "Mode")

        # Format columns
        tree.column("#0", width=0, stretch=tk.NO)  # First column is the phantom column
        tree.column("Condition", anchor=tk.W, width=150)
        tree.column("Mean", anchor=tk.W, width=100)
        tree.column("Median", anchor=tk.W, width=100)
        tree.column("Mode", anchor=tk.W, width=100)

        # Define headings
        tree.heading("#0", text="", anchor=tk.W)
        tree.heading("Condition", text="Condition", anchor=tk.W)
        tree.heading("Mean", text="Mean", anchor=tk.W)
        tree.heading("Median", text="Median", anchor=tk.W)
        tree.heading("Mode", text="Mode", anchor=tk.W)

        # Insert data into treeview
        tree.insert(parent='', index='end', iid=0, text="",
                    values=("Site Height > 75", mean_value_siteHeight, median_value_siteHeight, mode_value_siteHeight))
        tree.insert(parent='', index='end', iid=1, text="",
                    values=("Date from 2001", mean_value_date, median_value_date, mode_value_date))

        # Add the treeview to the canvas using create_window
        self.canvasMMM.create_window((0, 0), window=tree, anchor='nw')

        # Update canvas scroll regions after adding widgets
        self.canvasMMM.config(scrollregion=self.canvasMMM.bbox("all"))

    # Contents of Tab 2
    def data_analysis_graph(self, tab):
        label = ttk.Label(tab, text="\nEIDs extracted:\nC118, C11, C18A\n")
        label.grid(row=3, column=0, padx=20, sticky="w")

        labelC = ttk.Label(tab, text="\nGraphical analysis\n")
        labelC.grid(row=4, column=0, padx=20, sticky="w")

        # Button to analyse extracted data for selected criteria
        graph = ttk.Button(tab, text="Analyse", command=lambda: self.display_graphs())
        graph.grid(row=6, column=0, padx=20, sticky="w")

        self.canvasGraph = tk.Canvas(tab, width=900, height=300, borderwidth=0)
        self.canvasGraph.grid(row=7, column=0, columnspan=8, padx=20, pady=20, sticky="nsew")

    # Displaying of three graphs of analysis
    def display_graphs(self):
        df = self.merged
        df_melted = df.melt(id_vars='EID',
                            value_vars=['Serv Label1', 'Serv Label2', 'Serv Label3', 'Serv Label4', 'Serv Label10'],
                            var_name='Service_Label', value_name='Label_Value')
        # Get counts for each EID - Label_Value combination
        df_grouped = df_melted.groupby(['EID', 'Label_Value']).size().unstack()

        # Plot stacked bar chart
        fig1 = plt.figure(figsize=(6, 2))
        df_grouped.plot(kind='bar', stacked=True, ax=fig1.add_subplot(111))
        plt.ylabel('Count')

        # Plotting distribution of frequency
        fig2 = plt.figure(figsize=(3, 2))
        sns.histplot(data=df, x="Freq.", kde=True, ax=fig2.add_subplot(111))
        plt.title('Frequency Distribution')

        # Plotting distribution of Block
        fig3 = plt.figure(figsize=(3, 2))
        sns.histplot(data=df, x="Block", kde=True, ax=fig3.add_subplot(111))
        plt.title('Block Distribution')

        # Embed graphs in Tab 2
        self.embed_in_tkinter(fig1,  0, 0, colspan=2)
        self.embed_in_tkinter(fig2, 1, 0)
        self.embed_in_tkinter(fig3, 1, 1)

    # Embedding the three graphs on Tab 2
    def embed_in_tkinter(self, fig, row, col, colspan = 1):
        canvas = FigureCanvasTkAgg(fig, self.canvasGraph)
        canvas.draw()
        canvas.get_tk_widget().grid(row=row, column=col, columnspan=colspan, sticky='nsew')

    # Contents of Tab 3
    def data_analysis_correlation(self, tab):
        label = ttk.Label(tab, text="\nEIDs extracted:\nC118, C11, C18A\n")
        label.grid(row=3, column=0, padx=20, sticky="w")

        labelC = ttk.Label(tab, text="\nCorrelation information\n")
        labelC.grid(row=4, column=0, padx=20, sticky="w")

        # Button to analyse extracted data for selected criteria
        correlation = ttk.Button(tab, text="Analyse", command=lambda: self.display_correlation())
        correlation.grid(row=6, column=0, padx=20, sticky="w")

        self.canvasCorr = tk.Canvas(tab, width=900, height=300, borderwidth=0)
        self.canvasCorr.grid(row=7, column=0, columnspan=8, padx=20, pady=20, sticky="nsew")

    # Creating correlation matrix and displaying it
    def display_correlation(self):
        df = self.merged
        labels_mapping = {label: idx for idx, label in enumerate(pd.concat(
            [df['Block'], df['Serv Label1'], df['Serv Label2'], df['Serv Label3'], df['Serv Label4'], df['Serv Label10']]).unique())}

        df['Block'] = df['Block'].map(labels_mapping)
        df['Serv Label1'] = df['Serv Label1'].map(labels_mapping)
        df['Serv Label2'] = df['Serv Label2'].map(labels_mapping)
        df['Serv Label3'] = df['Serv Label3'].map(labels_mapping)
        df['Serv Label4'] = df['Serv Label4'].map(labels_mapping)
        df['Serv Label10'] = df['Serv Label10'].map(labels_mapping)

        correlation_matrix = df[
            ['Freq.', 'Block', 'Serv Label1', 'Serv Label2', 'Serv Label3', 'Serv Label4', 'Serv Label10']].corr()
        fig, ax = plt.subplots()
        sns.heatmap(correlation_matrix, annot=True, cmap='coolwarm', ax=ax)

        canvas = FigureCanvasTkAgg(fig, self.canvasCorr)
        canvas.draw()
        canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=0)

    # Select CSV files for processing and analysis
    def select_file(self,file_entry):
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")],initialdir=os.curdir)
        file_entry.configure(state="normal")
        file_entry.delete(0, tk.END)
        file_entry.insert(tk.END, file_path)
        file_entry.configure(state="readonly")
        print(f"Selected file: {file_path}")
        self.text_area1.insert(tk.END, f"\nSelected file loaded: {file_path}")

    def display_recordsAnt_with_validation(self, filepath):
        if not filepath:
            messagebox.showwarning("Warning", "Please select a file first!")
            return
        self.display_recordsAnt(filepath)

    def clean_recordsAnt_with_validation(self, filepath):
        if not filepath:
            messagebox.showwarning("Warning", "Please select a file first!")
            return
        self.clean_recordsAnt(filepath)

    # Display dataframe
    def display_recordsAnt(self, file_path):
        df = pd.read_csv(file_path)
        self.pt = Table(self.canvas1, dataframe=df)
        self.pt.show()

    # Display dataframe
    def display_recordsParam(self, file_path):
        df = pd.read_csv(file_path, encoding='ISO-8859-1',)
        self.pt = Table(self.canvas1, dataframe=df)
        self.pt.show()

    # Cleaning of TXParamsDAB file
    def clean_recordsParam(self, file_path):
        # Read in file and check for empty data
        # Issue with encoding of the Params csv
        df = pd.read_csv(file_path, encoding='ISO-8859-1', na_values=' ')
        df['Multiplex'] = df['EID'].str.extract('(C18A|C18F|C188)')
        df_filtered = df[df['Multiplex'].notnull()]
        df_filtered.to_csv('Filtered_Params.csv', index=False)
        self.text_area1.insert(tk.END, f"\nTXParam file cleaned")
        self.cleaned_Params = df_filtered
        self.tab1_action_completed1 = True

    # Cleaning of TXAntennaDAB file
    def clean_recordsAnt(self, file_path):
        # Read in file and check for empty data
        df = pd.read_csv(file_path, na_values=' ')
        ngrs = ['NZ02553847', 'SE213515', 'NT05399374', 'NT252675908']
        df_filtered = df[~df['NGR'].isin(ngrs)]
        df_filtered.to_csv('Filtered_Antenna.csv', index=False)
        self.text_area1.insert(tk.END, f"\nTXAntenna file cleaned")
        self.cleaned_Ants = df_filtered
        self.tab1_action_completed2 = True

    def save_file(self, file_path):
        self.text_area1.insert(tk.END, f"\nFile saved")

    # Clean and extract data from TXParamDAB and TXAntennaDAB to merged dataframe for anaylss
    def extract_data_frames(self):
        df_merged = pd.merge(self.cleaned_Params, self.cleaned_Ants, on="id")
        df_merged = df_merged.rename(columns={'In-Use Ae Ht': 'Aerial height(m)', 'In-Use ERP Total': 'Power(kW)'})
        # Clean format of Power(kW) number
        df_merged['Power(kW)'] = df_merged['Power(kW)'].str.replace('.', '').str.replace(',', '.').astype(float)
        # Column names had a space after them
        df_merged.columns = df_merged.columns.str.strip()
        df_merged = df_merged[
            ['id', 'Date', 'Ensemble', 'EID', 'Site', 'Site Height', 'Freq.', 'Block', 'Aerial height(m)', 'Power(kW)',
             'Serv Label1', 'Serv Label2', 'Serv Label3', 'Serv Label4', 'Serv Label10', 'Lat', 'Long']]
        # Date format needs to be cleaned
        df_merged['Date'] = pd.to_datetime(df_merged['Date'], format='%d/%m/%Y')

        self.text_area2.insert(tk.END, f"\nCleaned data from TXParam and TXAntenna has now been merged into a single datframe based on the criteria below.\nSee displayed records below.")
        df_merged.to_csv('merged.csv', index=False)
        self.merged = df_merged

        self.pt = Table(self.canvas2, dataframe=df_merged)
        self.pt.show()

    # Save cleaned and extracted data (merge) to MongoDB collection
    def save_extracted_data(self):
        client = pymongo.MongoClient("mongodb://localhost:27017/")
        db = client["APSummative"]
        collection = db["merged"]

        # Convert the DataFrame to a list of dictionaries and insert it into the collection
        records = self.merged.to_dict(orient='records')
        collection.insert_many(records)
        self.text_area2.insert(tk.END, f"\nData added to Mongo DB - Database: APSummative, Collection: merged")

    def tab_changed(self, event):
        # If the second tab is selected and the data clean isn't completed for both files, show a warning
        if (self.notebook.index("current") == 1 or self.notebook.index("current") == 2) and not self.tab1_action_completed1 or not self.tab1_action_completed2:
            messagebox.showwarning("Warning", "Please complete the Loading and Cleaning of both files first!")
            self.notebook.select(0)

# Mainline starts here
if __name__ == "__main__":
    app = TabbedInterface()
    app.mainloop()
