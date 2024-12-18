
from apify_client.client import ApifyClient
import pandas as pd
pd.set_option('display.max_colwidth', None)
import json
import configparser

# Initialize the ApifyClient with your API token
print("Packages installed successfully and client initialised!")


def getConfig():
    config = configparser.ConfigParser()
    config.read('config.ini') 
    return config

def readProjectListFile(path, sheetName = 'Results', colName = 'Name'):
    df = pd.read_excel(path, sheet_name=sheetName)
    projects = []
    for p in df[colName]:
        projects.append(p)
    return projects

def runQuery(client, q):
    results = []
    outputs = []
    # Prepare the Actor input JSON
    run_input = {
        "queries": q,
        "maxPagesPerQuery": 1,
        "resultsPerPage": 100,
        "maxConcurrency": 1,
        "customDataFunction": """async ({ input, $, request, response, html }) => {
      return {
        pageTitle: $('title').text(),
      };
    };""",
    }
    
    print("Generated a query for " + q)
    
    # Run the Actor and wait for it to finish
    run = client.actor("apify/google-search-scraper").call(run_input=run_input)
    
    # Fetch and print Actor results from the run's dataset (if there are any)
    for item in client.dataset(run["defaultDatasetId"]).iterate_items(clean=True):
        results = item.get('organicResults')
        results.append(item.get('paidResults'))
        
    for r in results:
        if(type(r) is dict):
            outputs.append([q, makeHyperlink(r.get('title'), r.get('url')), r.get('url')])
        
    print("Finished query " + q)
    return outputs
        
def makeHyperlink(text, url):
    return '=HYPERLINK("%s", "%s")' % (url, text)

def setHardCodedQueries():
    queries = []
    query = 'Project Reignite: Turning Farm Waste to Climate Action'
    queries.append(query)
    query = 'Jinshan Songlin Market Swine Manure Utilization Project'
    queries.append(query)
    return queries

def main():
    config = getConfig()
    client = ApifyClient(config['Main']['api key'])
    outputs = []
    
    projects = readProjectListFile(path = config['Main']['project list file'], 
                                   sheetName = config['Main']['sheet name'], 
                                   colName = config['Main']['column name'])
    
    for project in projects:
        outputs = outputs + (runQuery(client, project))
        
    df = pd.DataFrame(outputs, columns=['Project Name', 'Page Title', 'URL'])
    df.to_excel('results.xlsx')

main()

print('Done')