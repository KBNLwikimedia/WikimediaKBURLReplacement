import mwclient
import re
import requests
import os
import logging
from typing import List, Set, Optional
from dotenv import load_dotenv
from tqdm import tqdm  # Import tqdm for progress bar
from zope.interface import named

"""
This script automates the process of updating Wikimedia Commons files using the Wikimedia Commons API.
It searches for files matching a specific query, checks if they contain a target pattern,
and replaces it with a new pattern if found. The updated files are then saved back to Commons.

Features:
- Searches for files matching the given criteria.
- Logs into Wikimedia Commons using a bot account.
- Checks and updates file descriptions where needed.
- Saves changes with an appropriate edit summary.
- Displays a progress bar while processing files.

Requirements:
- Install `mwclient`, `requests`, `python-dotenv`, and `tqdm` via pip if not already installed.
- Ensure you have a Wikimedia bot account with appropriate permissions.
- Store credentials securely in a `.env` file.

Created by Olaf Janssen, Wikimedia coordinator of KB, national library of the Netherlands 
with much help from ChatGPT. 

Latest update: 16 September 2025

License = CC0, public domain.

"""

# Configure logging
logging.basicConfig(
    format="%(asctime)s - %(levelname)s - %(message)s",
    level=logging.INFO
)

# Load environment variables from a .env file
load_dotenv()

# Wikimedia Commons login credentials (read from environment variables)
USERNAME = os.getenv("WIKIMEDIA_USERNAME", "").strip()
PASSWORD = os.getenv("WIKIMEDIA_PASSWORD", "").strip()
USER_AGENT = os.getenv("WIKIMEDIA_USER_AGENT", "").strip()

if not USERNAME or not PASSWORD:
    logging.error(
        "Wikimedia credentials are missing. Please set WIKIMEDIA_USERNAME and WIKIMEDIA_PASSWORD in a .env file or as environment variables.")
    raise SystemExit(1)

if not USER_AGENT:
    logging.warning("User-Agent is not set. Some APIs may reject requests without it.")

# Wikimedia API URL and search parameters
SEARCH_URL = "https://commons.wikimedia.org/w/api.php"
SEARCH_PARAMS = {
    "action": "query",
    "format": "json",
    "list": "search",
    "srsearch": "insource:\"Category:Media from Delpher\"",
    "srlimit": 500,  # Adjust limit as needed, max=500 per query
    "srnamespace": 6,  # Namespace 6 = Files
    "sroffset": 0
}


# Define text patterns to search and replace
# OLD_PATTERN = re.compile(r"\}\}\n== \{\{int:license-header\}\} ==\n\{\{PD-old-70-expired\}\}", re.MULTILINE)
# NEW_PATTERN = r"}}\n{{Delpher}}\n== {{int:license-header}} ==\n{{PD-old-70-expired}}"

# OLD_PATTERN = re.compile(r"\}\}\n\n\[\[Category:De katholieke encyclopaedie\]\]", re.MULTILINE)
# NEW_PATTERN = r"}}\n{{Delpher}}\n[[Category:De katholieke encyclopaedie]]"

# OLD_PATTERN = re.compile(r"\}\}\n\[\[Category:", re.MULTILINE)
# NEW_PATTERN = r"}}\n{{Delpher}}\n[[Category:"

OLD_PATTERN = re.compile(r"\[\[Category:Media from Delpher\]\]", re.MULTILINE)
NEW_PATTERN = r""


# Number of files to process per run
MAX_FILES = 60000
edit_summary = "Removed Category:Media from Delpher"


def get_files() -> List[str]:
    """
    Retrieve up to MAX_FILES unique file titles from Wikimedia Commons based on a search query.
    Ensures that only MAX_FILES files are collected even if srlimit is higher.

    Returns:
        List[str]: A list of unique file titles.
    """
    all_files: Set[str] = set()
    params = SEARCH_PARAMS.copy()
    session = requests.Session()
    try:
        while len(all_files) < MAX_FILES:
            params["srlimit"] = min(SEARCH_PARAMS["srlimit"], MAX_FILES - len(all_files))  # Adjust limit dynamically
            response = session.get(SEARCH_URL, params=params, headers={"User-Agent": USER_AGENT})
            response.raise_for_status()
            data = response.json()

            # Extract file titles and ensure uniqueness
            files = {page["title"] for page in data.get("query", {}).get("search", [])}
            all_files.update(files)

            logging.info(f"Retrieved {len(all_files)} unique files so far...")

            # Stop if the required number of files is reached or if no more results are available
            if len(all_files) >= MAX_FILES or "continue" not in data:
                break

            # Update the continue parameter for the next request
            params.update(data["continue"])

    except requests.RequestException as e:
        logging.error(f"Error fetching file list: {e}")
        return []

    logging.info(f"Final count: {len(all_files)} unique files retrieved.")
    return list(all_files)


def login_to_commons() -> Optional[mwclient.Site]:
    """
    Log in to Wikimedia Commons using mwclient.

    Returns:
        mwclient.Site: Authenticated Wikimedia Commons site instance or None if login fails.
    """
    try:
        site = mwclient.Site("commons.wikimedia.org", clients_useragent=USER_AGENT)
        site.login(USERNAME, PASSWORD)
        logging.info("Successfully logged in to Wikimedia Commons.")
        return site
    except mwclient.LoginError as e:
        logging.error(f"Login failed: {e}")
    except Exception as e:
        logging.error(f"Unexpected error during login: {e}")

    return None


def has_multiple_delpher(text: str, limit: int) -> bool:
    """
    Checks if the text contains the {{Delpher}} template more than 'limit' times.

    Args:
        text (str): The text to check.

    Returns:
        bool: True if {{Delpher}} appears more than 'limit' times. False otherwise.
    """
    return text.count("{{Delpher}}") > int(limit)


def process_files(site: mwclient.Site, file_titles: List[str], edit_summary: str) -> None:
    """
    Check each file's wikitext and replace the target pattern if found,
    ensuring that {{Delpher}} is not already present. A progress bar is displayed.

    Args:
        site (mwclient.Site): Authenticated Wikimedia Commons site instance.
        file_titles (List[str]): List of file titles to process.
    """
    if not file_titles:
        logging.warning("No files to process.")
        return

    for count, title in enumerate(tqdm(file_titles, desc="Processing files", unit="file"), start=1):
        try:
            page = site.pages[title]
            text = page.text()

            # Skip if {{Delpher}} is already present before the license header
            if has_multiple_delpher(text, 2):  # text contains {{Delpher}} once or more
                logging.info(f"Skipping {title} - Template {{{{Delpher}}}} already present once or more times.")
                continue

            # Replace only if the exact OLD_PATTERN is found
            if OLD_PATTERN.search(text):
                new_text = OLD_PATTERN.sub(NEW_PATTERN, text)
                page.save(new_text, summary=edit_summary)
                logging.info(f"{count} - Updated: {title}")
            else:
                logging.info(f"Pattern not found, skipping: {title}")
        except mwclient.errors.APIError as e:
            logging.error(f"API error processing {title}: {e}")
        except Exception as e:
            logging.error(f"Error processing {title}: {e}")


def main() -> None:
    """
    Main function to orchestrate file retrieval, login, and processing.
    """
    file_titles = get_files()
    if not file_titles:
        logging.error("No files found or error retrieving files.")
        return

    site = login_to_commons()
    if not site:
        logging.error("Login failed, exiting script.")
        return

    #process_files(site, file_titles, edit_summary)
    #logging.info(f"Finished processing {len(file_titles)} files.")


if __name__ == "__main__":
    main()
