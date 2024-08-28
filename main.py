import os
import tkinter as tk
from tkinter import filedialog, Text, ttk
import pandas as pd
import matplotlib.pyplot as plt

# Function to select folder and list files with index and type-based coloring
def select_folder():
    folder_path = filedialog.askdirectory()
    if folder_path:
        files_list.delete(1.0, tk.END)  # Clear the text area
        files = os.listdir(folder_path)
        files.sort()  # Sort files alphabetically

        for index, file in enumerate(files, start=1):
            file_extension = os.path.splitext(file)[1].lower()  # Get file extension
            if file_extension in ['.txt', '.md']:
                color = 'blue'
            elif file_extension in ['.jpg', '.png', '.gif']:
                color = 'green'
            elif file_extension in ['.pdf']:
                color = 'red'
            else:
                color = 'black'

            files_list.insert(tk.END, f"{index}. {file}\n")
            files_list.tag_add(str(index), f"{index}.0", f"{index}.end")
            files_list.tag_config(str(index), foreground=color)

# Function to select and read an Excel file, then display the data
def read_excel():
    global df
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        excel_text.delete(1.0, tk.END)  # Clear the text area
        df = pd.read_excel(file_path)  # Read Excel file
        excel_text.insert(tk.END, df.to_string(index=False))  # Display data

# Function to plot the data from the Excel file
def plot_data():
    if df is not None:
        df.plot(kind='line')  # Plot the data (you can change the plot type)
        plt.show()  # Display the plot

# Function to display a statistical summary of the Excel data
def show_statistics():
    if df is not None:
        stats = df.describe()  # Get statistical summary
        excel_text.delete(1.0, tk.END)  # Clear the text area
        excel_text.insert(tk.END, stats.to_string())  # Display statistical summary

# Create the main window
root = tk.Tk()
root.title("Multi-Function Application")

# Create a Notebook (tabbed interface)
notebook = ttk.Notebook(root)
notebook.pack(fill=tk.BOTH, expand=True)

# Create the first tab for File Explorer
tab1 = ttk.Frame(notebook)
notebook.add(tab1, text="File Explorer")

# Add the file explorer components to the first tab
select_button = tk.Button(tab1, text="Select Folder", padx=20, pady=10, command=select_folder)
select_button.pack()
files_list = Text(tab1, wrap=tk.WORD, padx=10, pady=10)
files_list.pack(fill=tk.BOTH, expand=True)

# Create the second tab for Excel Reader and Analytics
tab2 = ttk.Frame(notebook)
notebook.add(tab2, text="Excel Reader")

# Add the Excel reader components to the second tab
excel_button = tk.Button(tab2, text="Select Excel File", padx=20, pady=10, command=read_excel)
excel_button.pack()

# Add the Plot button
plot_button = tk.Button(tab2, text="Plot Data", padx=20, pady=10, command=plot_data)
plot_button.pack()

# Add the Show Statistics button
stats_button = tk.Button(tab2, text="Show Statistics", padx=20, pady=10, command=show_statistics)
stats_button.pack()

# Add a text area to display Excel data and statistics
excel_text = Text(tab2, wrap=tk.WORD, padx=10, pady=10)
excel_text.pack(fill=tk.BOTH, expand=True)

# Start the GUI loop
root.mainloop()
