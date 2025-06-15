<h1>
    <p align="center">Money Manager Script</p>
</h1>

This Python script is designed to retrieve and prep financial data for my custom Excel Money Manager tool which I use for managing financial transactions, investments, and eStatements. Tasks that are automated include retrieving transaction data from online banking platforms, categorizing transactions, fetching investment information from Robinhood and Coinbase, and downloading eStatements from online banking portals.

<div style="text-align: center;">
    <img src="Money%20Manager%20Pic.webp" alt="Money Manager Logo" width="500"/>
</div>

## Requirements

- Python 3.x
- Required Python packages:
    - `dateutil`
    - `coinbase`
    - `robin_stocks`
    - `selenium`
    - `xlwings`
    - `pandas`
    - `PyPDF2`

## Usage

To use the script, follow these steps:

1. Ensure that all required Python packages are installed.
2. Set up the necessary configurations in the Excel workbook referenced by the script.
3. Run the script using Python.
   - If ran from the terminal (rather than from Excel VBA), run with the current directory being a subfolder of the folder holding the .xlsm workbook. And be sure to manually pass along the creds while instantiating the object. 
     - `from retrieve_creds import retrieve_creds_for_money_manager`
     - `from Money_Manager import Money_Manager`
     - `creds = retrieve_creds_for_money_manager()`
     - `money_manager = Money_Manager(creds)`
   - The creds are actually optional and not needed for calling methods that don't require API access

The script performs various tasks such as:

- Retrieving transaction data from online banking portals.
- Categorizing transactions and updating Excel spreadsheets accordingly.
- Fetching investment information from Robinhood and Coinbase.
- Downloading eStatements from online banking portals.

### The investment portfolio part:

| Symbol | Name | Investment Type | Sector | Industry | Current Quantity | Current Equity | All Time Net Loss or Gain |
|--------|------|-----------------|--------|----------|------------------|----------------|---------------------------|
|        |      |                 |        |          |                  |                |                           |

## Configuration

Before running the script, make sure to set up the following configurations:

- Excel Workbook: The script requires access to an Excel workbook containing necessary reference data and configurations. Update the file path in the script to point to the correct workbook.
- Account Credentials: Provide credentials for accessing online banking platforms, Robinhood, and Coinbase if required.
- Browser Driver: Ensure that the appropriate browser driver (e.g., ChromeDriver) is installed and its path is specified correctly in the script.

<p align="right">Click <a href="https://github.com/bhyman67/Functionalities-for-my-Money-Manager">here</a> to view the code in this project's repository<p>
