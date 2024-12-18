from apify_client import ApifyClient
import pandas as pd
import configparser
import tkinter as tk
from tkinter import filedialog, ttk
import threading
import requests
import re

def getConfig():
    config = configparser.ConfigParser()
    config.read('config.ini') 
    return config

def readProjectListFile(path, sheetName='Results', colName='Name', extraCols=[]):
    df = pd.read_excel(path, sheet_name=sheetName)
    projects = []
    
    for _, row in df.iterrows():
        # Start with the main search term
        project_name = row[colName]  # Just the project name
        query = project_name
        
        # Add additional terms (if provided)
        for col in extraCols:
            if not pd.isna(row[col]):  # Only add non-empty values
                query += f" {row[col]}"
        
        projects.append((project_name, query))
    
    return projects


def runQuery(client, project_name, q, progress_bar, progress_step, log_text_widget):
    results = []
    outputs = []

    # Define additional terms to refine your search
    additional_terms = "carbon credit verification"
    refined_query = f"{q} {additional_terms}"

    # Prepare the Actor input JSON
    run_input = {
        "queries": refined_query,
        "maxPagesPerQuery": 1,
        "resultsPerPage": 50,
        "maxConcurrency": 1,
        "customDataFunction": """async ({ input, $, request, response, html }) => {
            return {
                pageTitle: $('title').text(),
                url: request.url
            };
        };""",
    }

    update_log_window(log_text_widget, f"Generated a query for {refined_query}")

    # Run the Actor and wait for it to finish
    run = client.actor("apify/google-search-scraper").call(run_input=run_input)
    
    # Fetch and process results
    for item in client.dataset(run["defaultDatasetId"]).iterate_items(clean=True):
        if 'organicResults' in item:
            results.extend(item['organicResults'])
        if 'paidResults' in item:
            results.extend(item['paidResults'])
    
    for r in results:
        if isinstance(r, dict):
            title = r.get('title', "No Title")  # Default to 'No Title' if title is missing
            url = r.get('url', "#")  # Default to '#' if URL is missing

            # Skip URLs that trigger file downloads
            if is_file_download(url):
                update_log_window(log_text_widget, f"Skipped download link: {url}")
                continue
            
            if title.strip() and url.strip():  # Ensure both title and URL are not empty
                outputs.append([project_name, makeHyperlink(title, url), url])

    # Update the progress bar
    progress_bar['value'] += progress_step

    update_log_window(log_text_widget, f"Finished query for project: {refined_query}")
    return outputs

def is_file_download(url):
    """Check if a URL triggers a file download by inspecting headers."""
    try:
        response = requests.head(url, allow_redirects=True, timeout=5)
        content_disposition = response.headers.get('Content-Disposition', '')
        
        # If Content-Disposition indicates attachment, it's likely a file download
        if 'attachment' in content_disposition.lower():
            return True
        
        # Optionally, check if the URL ends with common file extensions
        if re.search(r'\.(zip|docx|xlsx|exe|tar|gz|rar|csv|pptx)(\?.*)?$', url, re.IGNORECASE):
            return True
    except requests.RequestException:
        # If there's an error (timeout, etc.), assume it's not a download
        return False
    
    return False

def makeHyperlink(text, url):
    """Generate a safe hyperlink formula for Excel."""
    if not text:  # Handle None or empty text
        text = "No Title"
    if not url:  # Handle None or empty URL
        url = "#"

    # Escape double quotes in text and URL
    text = str(text).replace('"', '""')
    url = str(url).replace('"', '""')

    # Return the hyperlink as a valid string
    return f'=HYPERLINK("{url}", "{text}")'

def main(file_path, progress_bar, log_window):
    config = getConfig()
    client = ApifyClient(config['Main']['api key'])
    outputs = []

   # Read projects with additional columns
    try:
        projects = readProjectListFile(
            path=file_path,
            sheetName='Results',
            colName='Name',
            extraCols=['Country/Area']  # Specify additional columns
        )
    except ValueError as e:
        update_log_window(log_window, str(e))
        return

    # Calculate progress step based on the total number of projects
    total_projects = len(projects)
    if total_projects == 0:
        update_log_window(log_window, "No projects found in the 'Name' column.")
        return

    progress_step = 100 / total_projects

     # Loop through each refined query and run the search
    for project_name, query in projects:
        update_log_window(log_window, f"Processing project: {project_name}")
        outputs.extend(runQuery(client, project_name, query, progress_bar, progress_step, log_window))

    # Ask the user where to save the output file
    save_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        title="Save as"
    )
    
    if save_path:  # Save the file if a location was provided
        df = pd.DataFrame(outputs, columns=['Project Name', 'Page Title', 'URL'])
        df.fillna("", inplace=True)  # Replace NaN with empty strings
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

def browse_file(progress_bar, log_text_widget):
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
    if file_path:
        progress_bar['value'] = 0  # Reset progress bar
        update_log_window(log_text_widget, f"Selected file: {file_path}")

        # Run the main function in a separate thread to keep the UI responsive
        threading.Thread(target=main, args=(file_path, progress_bar, log_text_widget)).start()
    else:
        update_log_window(log_text_widget, "No file selected.")

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
