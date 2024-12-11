import os
import requests
from docx import Document
from dotenv import load_dotenv
import json

# Load environment variables
load_dotenv()

# Your Notion token and list of databases
NOTION_TOKEN = os.getenv("NOTION_TOKEN")
TARGET_DATABASES = json.loads(os.getenv("DATABASES"))

# Notion API base URL
NOTION_API_URL = "https://api.notion.com/v1/databases/{}/query"

# Headers for authorization and content type
HEADERS = {
    "Authorization": f"Bearer {NOTION_TOKEN}",
    "Content-Type": "application/json",
    "Notion-Version": "2022-06-28"
}

# Function to query a Notion database with pagination
def query_notion_database(database_id, start_cursor=None):
    url = NOTION_API_URL.format(database_id)
    data = {}
    if start_cursor:
        data["start_cursor"] = start_cursor

    response = requests.post(url, headers=HEADERS, json=data)
    if response.status_code == 200:
        return response.json()
    else:
        print(f"Failed to query Notion for database {database_id}: {response.status_code}")
        return None

# Function to determine the short description based on the customer name and database
def get_short_description(database_name, customer_name):
    if database_name == 'Digiküpsuse hindamine':
        if 'DMA T1' in customer_name:
            return 'DMA (Digital Maturity Assessment) T1'
        elif 'DMA T0' in customer_name:
            return 'DMA (Digital Maturity Assessment) T0'
    elif database_name == 'AI nõustamine':
        if 'AI Nõustamine' in customer_name:
            number = customer_name.split()[-1]
            return f'AI suitability analysis {number}'
    elif database_name == 'Finantseerimise nõustamine – avalikud meetmed':
        if 'Avalikud meetmed' in customer_name:
            number = customer_name.split()[-1]
            return f'Support to find funding – public measures {number}'
    elif database_name == 'Finantseerimise nõustamine – erakapitali kaasamine':
        if 'Erakapitali kaasamine' in customer_name:
            number = customer_name.split()[-1]
            return f'Consulting to find funding – private investments {number}'
    elif database_name == 'Robotiseerimise nõustamine':
        if 'Robotiseerimise nõustamine' in customer_name:
            number = customer_name.split()[-1]
            return f'Robotics suitability analysis {number}'
    return ''

# Function to get service price based on database name
def get_service_price(database_name):
    prices = {
        'Digiküpsuse hindamine': '600',
        'AI nõustamine': '1400',
        'Robotiseerimise nõustamine': '1400',
        'Finantseerimise nõustamine – avalikud meetmed': '800',
        'Finantseerimise nõustamine – erakapitali kaasamine': '800'
    }
    return prices.get(database_name, '')

# Function to find projects in a specific database
def find_projects_in_database(database_id, database_name):
    projects = []
    start_cursor = None
    has_more = True

    # Iterate through all pages of results
    while has_more:
        data = query_notion_database(database_id, start_cursor)
        if not data:
            break

        for result in data["results"]:
            # Extract relevant properties
            properties = result["properties"]
            service_status = properties.get("Service status", {}).get("status", {}).get("name", "")
            edih_status = properties.get("EDIH platvormile sisestatud – Finalised", {}).get("select", None)

            # Safely extract project name if it exists
            project_name = ""
            projekt_field = properties.get("Projekt", {}).get("title", [])
            if projekt_field:  # Check if the list is not empty
                project_name = projekt_field[0].get("text", {}).get("content", "")

            # Extract the finishing date from 'Automaatne väli, VTA väljamakse tehtud (kpv)'
            finishing_date_field = properties.get("Automaatne väli, VTA väljamakse tehtud (kpv)")
            finishing_date = ""
            if finishing_date_field and finishing_date_field.get("date"):
                finishing_date = finishing_date_field.get("date", {}).get("start", "")

            # Extract the starting date based on the database
            if database_name == 'Digiküpsuse hindamine':
                starting_date_field = properties.get("Automaatne väli, DMA link ettevõttele saadetud (teenuse algus)")
                starting_date = ""
                if starting_date_field and starting_date_field.get("date"):
                    starting_date = starting_date_field.get("date", {}).get("start", "")
            else:
                # For other databases, assuming 'Esmanõustamise kuupäev' is a text field
                starting_date_field = properties.get("Esmanõustamise kuupäev (ev külastuse kpv, teenuse osutamise alguse kpv)")
                starting_date = ""
                if starting_date_field and starting_date_field.get("rich_text"):
                    starting_date = starting_date_field.get("rich_text", [{}])[0].get("text", {}).get("content", "")

            # Generate the short description based on the project name
            short_description = get_short_description(database_name, project_name)

            # Check conditions: Service status == 'Finalised' and EDIH platvormile sisestatud == None
            if service_status == "Finalised" and edih_status is None:
                projects.append((project_name, short_description, starting_date, finishing_date))

        # Check if there are more pages to fetch
        start_cursor = data.get("next_cursor", None)
        has_more = data.get("has_more", False)

    return projects

# Function to save results to a Word document with separate tables
def save_to_word(all_projects):
    document = Document()
    document.add_heading('Projects to Update', 0)

    # For each service (database), create a new table
    for database_name, projects in all_projects.items():
        if projects:
            document.add_heading(database_name, level=1)

            # Create a table with 9 columns
            table = document.add_table(rows=1, cols=9)
            table.style = 'Table Grid'

            # Define column headers
            headers = [
                'Customer', 'Service category delivered', 'Short description of the service',
                'Technology type used', 'Service price', 'Price invoiced to customer',
                'Starting date', 'Finishing date (expected)', 'Service Status'
            ]
            
            # Add headers to the table
            hdr_cells = table.rows[0].cells
            for i, header in enumerate(headers):
                hdr_cells[i].text = header

            # Add projects to the table
            for project_name, short_description, starting_date, finishing_date in projects:
                row_cells = table.add_row().cells
                row_cells[0].text = project_name  # Customer (Projekt)

                # For 'Finantseerimise' categories, set specific Service category
                if database_name in ['Finantseerimise nõustamine – avalikud meetmed', 'Finantseerimise nõustamine – erakapitali kaasamine']:
                    row_cells[1].text = 'Support to find investment'
                else:
                    row_cells[1].text = 'Test before invest'  # Service category delivered

                row_cells[2].text = short_description  # Short description of the service
                row_cells[3].text = 'AI'  # Technology type used
                row_cells[4].text = get_service_price(database_name)  # Service price
                row_cells[5].text = '0'  # Price invoiced to customer
                row_cells[6].text = starting_date  # Starting date
                row_cells[7].text = finishing_date  # Finishing date (expected)
                row_cells[8].text = 'Finalised'  # Service Status

    # Save the document to a file
    document_path = "projects_to_update_with_tables.docx"
    document.save(document_path)
    print(f"Document saved as {document_path}")

# Function to check all databases
def check_all_databases():
    all_projects_to_update = {}

    # Iterate over each database
    for database_id, database_name in TARGET_DATABASES.items():
        print(f"Checking database: {database_name} (ID: {database_id})")
        projects_to_update = find_projects_in_database(database_id, database_name)

        if projects_to_update:
            all_projects_to_update[database_name] = projects_to_update
            print(f"Projects to update in {database_name}:")
            for project, description, start_date, end_date in projects_to_update:
                print(f"- {project} ({description}) - Start date: {start_date}, Finishing date: {end_date}")
        else:
            print(f"No projects found that require updates in {database_name}.")

    return all_projects_to_update

if __name__ == "__main__":
    all_projects = check_all_databases()

    if all_projects:
        save_to_word(all_projects)
    else:
        print("No projects found that require updates in any database.")
