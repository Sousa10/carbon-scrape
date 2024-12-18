from apify_client import ApifyClient 
import pandas as pd
import configparser
import tkinter as tk
from tkinter import filedialog, ttk
import logging
import threading

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def getConfig():
    config = configparser.ConfigParser()
    config.read('config.ini') 
    return config

def readProjectListFile(path, sheetName='Results', colName='Name'):
    df = pd.read_excel(path, sheet_name=sheetName)
    return df[colName].tolist()

# Your existing runQuery function...
def runQuery(client, q, progress_bar, progress_step, log_text_widget):
    results = []
    outputs = []

    # Define additional terms to refine your search
    additional_terms = "carbon credit verification site:.org OR site:.gov"
    refined_query = f"{q} {additional_terms}"

    run_input = {
        "queries": refined_query,
        "maxPagesPerQuery": 1,
        "resultsPerPage": 50,
        "maxConcurrency": 1,
        "customDataFunction": """async ({ input, $, request, response, html }) => {
                return {
                    pageTitle: $('title').text(),
                };
            };""",
    }

    update_log_window(log_text_widget, f"Running query for project: {q}")

    run = client.actor("apify/google-search-scraper").call(run_input=run_input)
    
    for item in client.dataset(run["defaultDatasetId"]).iterate_items(clean=True):
        if 'organicResults' in item:
            results.extend(item['organicResults'])
        if 'paidResults' in item:
            results.extend(item['paidResults'])
    
    for r in results:
        if isinstance(r, dict) and not any(bad_word in r.get('url') for bad_word in ['badwebsite.com', 'spam.com']):
            outputs.append([q, f'=HYPERLINK("{r.get("url")}", "{r.get("title")}")', r.get('url')])

    # Update the progress bar and log text
    progress_bar['value'] += progress_step
    progress_bar.update_idletasks()
    update_log_window(log_text_widget, f"Finished query for project: {q}")

    return outputs

def main(file_path, progress_bar, log_window):
    config = getConfig()
    client = ApifyClient(config['Main']['api key'])
    outputs = []

    # Read all rows from the "Name" column in the Excel file
    projects = readProjectListFile(path=file_path,
                                   sheetName='Results',  # Adjust sheet name as needed
                                   colName='Name')  # Adjust column name as needed
    
    # Debug: Print all projects for verification
    update_log_window(log_window, f"Projects found: {projects}")
    
    total_projects = len(projects)
    if total_projects == 0:
        update_log_window(log_window, "No projects found in the 'Name' column.")
        return

    # Calculate progress step based on the total number of projects
    progress_step = 100 / total_projects

    # Loop through each project and run the query
    for project in projects:
        update_log_window(log_window, f"Processing project: {project}")
        # project_results = runQuery(client, project, progress_bar, progress_step, log_window)
        # outputs.extend(project_results)
        outputs = outputs + (runQuery(client, project, progress_bar, progress_step, log_window))

        # Debug: Ensure the outputs list is populated correctly

    update_log_window(log_window, f"Total rows in output: {len(outputs)}")

    # Ask the user where to save the output file
    save_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        title="Save as"
    )
    
    if save_path:  # Save the file if a location was provided
        df = pd.DataFrame(outputs, columns=['Project Name', 'Page Title', 'URL'])
        df.to_excel(save_path, index=False)
        update_log_window(log_window, f"File saved as: {save_path}")
    else:
        update_log_window(log_window, "File save canceled")

    update_log_window(log_window, "Process complete")


def create_log_window(root):
    log_window = tk.Text(root, wrap='word', state='disabled', height=10)
    log_window.pack(pady=10)
    return log_window

def update_log_window(log_window, message):
    log_window.config(state='normal')
    log_window.insert('end', message + '\n')
    log_window.see('end')
    log_window.config(state='disabled')
    log_window.update_idletasks()

# Function to browse and select the file
def browse_file(progress_bar, log_text_widget):
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
    if file_path:
        progress_bar['value'] = 0  # Reset progress bar
        update_log_window(log_text_widget, f"Selected file: {file_path}")

        # Run the main function in a separate thread to keep the UI responsive
        threading.Thread(target=main, args=(file_path, progress_bar, log_text_widget)).start()
    else:
        update_log_window(log_text_widget, "No file selected.")

# Function to create the GUI
def create_gui():
    root = tk.Tk()
    root.title("Excel File Selector")

    # Create and pack widgets
    progress_bar = ttk.Progressbar(root, length=400, mode='determinate')
    progress_bar.pack(pady=10)

    log_text_widget = create_log_window(root)

    browse_button = tk.Button(root, text="Select Excel File", command=lambda: browse_file(progress_bar, log_text_widget))
    browse_button.pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    create_gui()