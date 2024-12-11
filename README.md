# Notion Project Update Script

This script interacts with the Notion API to retrieve, process, and organize project data from multiple Notion databases. It identifies projects that require updates and generates a structured Word document with the results.

---

## Table of Contents

1. [Features](#features)
2. [Technologies Used](#technologies-used)
3. [Installation](#installation)
4. [Configuration](#configuration)
5. [Usage](#usage)
6. [Project Structure](#project-structure)

---

## Features

- **Fetch Data from Notion:** Queries Notion databases using the Notion API.
- **Filter and Process Projects:** Identifies projects marked as "Finalised" and not uploaded to the EDIH platform.
- **Generate Reports:** Outputs data into a Word document with tables for each database.
- **Customizable:** Easily modify database IDs and processing logic.

---

## Technologies Used

- **Languages:** Python
- **Libraries:**
  - `requests`: For API calls to Notion.
  - `python-docx`: For generating Word documents.
  - `dotenv`: For managing environment variables.
  - `json`: For processing JSON data.
- **APIs:** Notion API

---

## Installation

1. Clone the repository:

   ```bash
   git clone https://github.com/your-username/notion_edih.git
   cd notion_edih
   ```

2. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

3. Create a `.env` file in the root directory (see [Configuration](#configuration)).

---

## Configuration

### Environment Variables

Create a `.env` file 

### File: `config.py`

- The script loads environment variables using the `dotenv` library.
- Ensure that your Notion token and databases are correctly set in the `.env` file.

---

## Usage

1. Run the script:

   ```bash
   python main.py
   ```

2. The script will query each database, process the results, and generate a Word document named `projects_to_update_with_tables.docx`.

3. Logs and results will be displayed in the console.

---

## Project Structure

```plaintext
notion-projects-script/
├── main.py              # Main script
├── config.py            # Environment variable loader
├── .env                 # Environment variables (not included in version control)
├── README.md            # Documentation
```
