from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from dateutil.relativedelta import relativedelta
from coinbase.wallet.client import Client
from coinbase.rest import RESTClient
from datetime import datetime, timedelta
import robin_stocks.robinhood as rh
from selenium import webdriver
import xlwings as xw
import pandas as pd
import traceback
import PyPDF2
import json
import time
import os
import re

def increment(match):

    number = int(match.group(1))
    return f'${number + 1}'

def remove_visa(txn_description):
    return re.sub(r'\bVISA \b', '', txn_description)

def assign_credit_debit_ind(amt):

    if amt >= 0:
        return("Credit")
    else:
        return("Debit")

def check_for_existing_pdf(file_dir):
  
    exists_a_pdf = False
  
    for itm in os.listdir(file_dir):
      if itm.endswith(".pdf"):
          exists_a_pdf = True
          break
    
    return exists_a_pdf

def PDFmerge(pdfs, output_pdf_name):
  
    # create the pdf file merger object
    pdfMerger = PyPDF2.PdfFileMerger()
 
    # append pdfs one by one
    for pdf in pdfs:
        pdfMerger.append(pdf)
 
    # write combined pdf to output_pdf_name pdf file
    with open(output_pdf_name, 'wb') as f:
        pdfMerger.write(f)

def extract_and_remove_date(description, post_date):

    full_pattern = r'\bON\s\d{2}-\d{2}\s\d{4}\b'
    date_pattern = r'\s\d{2}-\d{2}'
    match = re.search(date_pattern, description)
    if match:
        date = match.group(0)
        description_without_date = re.sub(full_pattern, '', description)
        return date, description_without_date.strip()
    else:
        # Format post_date as ' MM-DD' (with leading space to match extracted format)
        try:
            formatted_date = post_date.strftime(' %m-%d')
        except Exception:
            formatted_date = str(post_date)  # fallback if not datetime-like
        return formatted_date, description

class Money_Manager:

    def __init__(self, creds = None):

        # If this class is being instantiated in the VBA source code (ran by the RunPython VBA funct)
        if __name__ == "Scripts.Money_Manager":

            self.wb = xw.Book.caller()

        else:

            self.wb = xw.Book("../Money Manager.xlsm")

        # Read in reference data from the money manager Excel workbook
        self.desc_cat_lookup = self.wb.sheets["Script Control Center & Ref Dta"].range("Table1").options(dict).value # dictionary
        self.desc_excludes = self.wb.sheets("Script Control Center & Ref Dta").range("Table2").value # list
        self.manual_desc = self.wb.sheets("Script Control Center & Ref Dta").range("Table3").options(pd.DataFrame, index = False, header = False).value # dataframe

        # Read in some script variables from the money manager Excel workbook
        self.account1_name = self.wb.sheets["Script Control Center & Ref Dta"].range("Account_1").value
        self.account2_name = self.wb.sheets["Script Control Center & Ref Dta"].range("Account_2").value
        self.account3_name = self.wb.sheets["Script Control Center & Ref Dta"].range("Account_3").value
        self.credit_card_account_name = self.wb.sheets["Script Control Center & Ref Dta"].range("Credit_Card_Account").value

        # Set credential variables if they were passed in
        if creds:
            self.firstbank_u = creds["FirstBank"][0]
            self.firstbank_p = creds["FirstBank"][1]
            self.robinhood_u = creds["Robinhood"][0]
            self.robinhood_p = creds["Robinhood"][1]
            self.coinbase_key_id = creds["Coinbase"][0]
            self.coinbase_key_secret = creds["Coinbase"][1]

    def __assign_exclude_ind(self, desc):

        # how to check if any items w/in a list are in a string
        truth_val = False
        if any(desc_exclude in desc for desc_exclude in self.desc_excludes):
            truth_val = True
        return(truth_val)

    def __categorize_description(self, desc):

        # loop through the dict
        for desc_substring in self.desc_cat_lookup:
            if desc_substring.upper() in desc.upper():
                # found a match
                return(self.desc_cat_lookup[desc_substring])

        # Return an empty string if no matches were found
        return("")

    def __del__(self):

        if __name__ != "Scripts.Money_Manager":

            self.wb.app.quit()

    def retrieve_data(self): # rename this function... 

        # You need to put the error handling back into this scraping routine... 

        # Grab what you need from Robinhood
        login = rh.authentication.login(self.robinhood_u, self.robinhood_p) 
        self.wb.sheets["Personal Investment Portfolio"].range( self.account3_name.replace(" ","_") ).value = rh.profiles.load_account_profile()["cash_available_for_withdrawal"]
        brokerage_interest_income_json_resp = rh.request_get("https://api.robinhood.com/accounts/sweeps")
        rh.authentication.logout()

        # Create a table out of the brokerage interest income data that you'll be adding to the transactions table
        brokerage_interest_income = pd.json_normalize( brokerage_interest_income_json_resp["results"] ) 
        brokerage_interest_income["Account"] = ' '.join(self.account3_name.split()[:2])
        brokerage_interest_income["Type"] = "INTEREST"
        brokerage_interest_income["Income_Expense_Exclude"] = False
        brokerage_interest_income = brokerage_interest_income[[
            "pay_date",
            "Account",
            "amount.amount",
            "reason",
            "Type",
            "direction",
            "Income_Expense_Exclude" 
        ]]

        brokerage_interest_income['direction'] = brokerage_interest_income['direction'].str.capitalize()

        brokerage_interest_income.rename(
            columns = {
                "pay_date":"Date",
                "amount.amount":"Amount",
                "reason":"Description",
                "direction":"Credit_Debit_Ind"
            },
            inplace = True
        )

        # Need to convert the date to just date without the time (with the format mm/dd/yyyy)
        # also only want to keep the data from November 1st, 2021 and on
        brokerage_interest_income["Date"] = pd.to_datetime(brokerage_interest_income["Date"])
        brokerage_interest_income['Date'] = brokerage_interest_income['Date'].dt.tz_localize(None)
        brokerage_interest_income = brokerage_interest_income[ brokerage_interest_income['Date'] >= pd.Timestamp('2021-11-01') ]
        #brokerage_interest_income = brokerage_interest_income[ brokerage_interest_income['Date'] >= '2021-11-01' ]
        brokerage_interest_income["Date"] = brokerage_interest_income["Date"].dt.strftime('%m/%d/%Y')

        # Set dates 
        crntDt = datetime.today().strftime('%m/%d/%Y')
        firstOfPrevMnth = (datetime.today() - relativedelta(months=1)).replace(day=1).strftime('%m/%d/%Y')

        # List all of your accounts
        accounts = []
        accounts.append( '{{accountType={account_name}, selectedNumber=2d83bcf05b214c9b1b032bef309d72b4}}'.format(account_name = self.account1_name) )
        accounts.append( '{{accountType={account_name}, selectedNumber=9e720c749c446ee65976669a391134fb}}'.format(account_name = self.account2_name) )
        accounts.append( '{{accountType={account_name}, selectedNumber=8c4a6dff17073338f88e3f5b3ae117a2}}'.format(account_name = self.credit_card_account_name) )

        # Instantiate the webdriver object
        service = webdriver.chrome.service.Service(self.wb.sheets["Script Control Center & Ref Dta"].range("Chromedriver").value)
        browser = webdriver.Chrome(service=service)
        browser.implicitly_wait(30)

        # Login to OB (is there a way to use credentials that are saved in the browser???)
        browser.get('https://www.efirstbank.com/')
        browser.find_element(By.ID, 'userId').send_keys(self.firstbank_u)
        browser.find_element(By.ID, 'password').send_keys(self.firstbank_p)
        browser.find_element(By.ID, 'logIn').click()

        # Grab account totals 
        # Current balance from account 1 
        time.sleep(10)
        browser.find_element(By.XPATH, '//*[@id="js-acct-name"]/span[1]')
        account1_current_balance = browser.find_element(By.XPATH, '//*[@id="js-ob-details-container"]/div/div/div[3]/div/div[2]/div[1]/ul/li[1]/strong/span').text
        # Click on account 2 and then grab the current balance from that
        browser.find_element(By.XPATH, '//*[@id="js-product-id-10620720"]/div[2]/div[1]/div/div[1]/p/span').click()
        time.sleep(3)
        account2_current_balance = browser.find_element(By.XPATH, '//*[@id="js-ob-details-container"]/div/div/div[3]/div/div[2]/div[1]/ul/li[1]/strong/span').text

        # Pull data for each account
        html_tables = []
        for account in enumerate(accounts):

            # Pull up the "Download Account Info" page

            # Click on the account tab
            browser.find_element(By.XPATH, '//*[@id="obTab"]/a').click()
            time.sleep(1)
            browser.find_element(By.LINK_TEXT, 'Downloads').click()

            # Select account
            browser.find_element(By.NAME, 'accountSelected').click()
            browser.find_element(By.XPATH, f"//option[@value = '{account[1]}']").click()

            # Set the date range (format is mm/dd/yyyy)
            browser.find_element(By.ID, 'dateRangeRadio').click()
            if account[0] == 0 or account[0] == 1:
                account = account[1].split(',')[0].split('=')[1].split(" ")[1]
            else:
                account = account[1].split(',')[0].split('=')[1]
            browser.find_element(By.NAME, 'fromDate').send_keys(firstOfPrevMnth)
            browser.find_element(By.NAME, 'toDate').send_keys(crntDt)

            # click  the view txns button
            browser.find_element(By.XPATH, "//input[@value='View Transactions']").click()

            # Find one of these two elements B4 the scrape to ensure that the page loads first
            elmnt_txt = browser.find_element(By.XPATH, "//table[@class='detail dataTable'] | //*[@id='contentContainer']/div[2]/div/p").text

            if "No transactions were found in the specified range." in elmnt_txt:

                pass 

            else:

                # Data Scrape
                html_table = pd.read_html(browser.page_source)[0]

                # Add an account col and then append data table to list
                html_table["Account"] = account
                html_table = html_table[["Date","Account","Amount","Description","Type"]]
                html_tables.append(html_table)

        # Combine all of the DFs and then export
        txns_df = pd.concat(html_tables)
        characters_to_replace = {
            "$":"",
            ",":"",
            "(":"-",
            ")":""
        }
        txns_df["Amount"] = txns_df["Amount"].replace(r'[\$,)]', '', regex=True).replace(r'[(]', '-', regex=True).astype(float) # txns_df["Amount"] = txns_df["Amount"].replace('[\$,)]', '', regex=True).replace('[(]','-',regex=True).astype(float)

        # Add a credit/debit indicator column and an income/expence exclude indicator column
        desc_excludes = self.wb.sheets("Script Control Center & Ref Dta").range("Table2").value
        # credit/debit indicator col
        txns_df["Credit_Debit_Ind"] = ""
        txns_df["Credit_Debit_Ind"] = txns_df["Amount"].apply(assign_credit_debit_ind)
        # indicator for transfers w/in internal accounts and credit card payments
        txns_df["Income_Expense_Exclude"] = ""
        txns_df["Income_Expense_Exclude"] = txns_df["Description"].apply(self.__assign_exclude_ind)

        # Log out and close both the browser and db cnxn
        time.sleep(2)
        browser.find_element(By.XPATH, "//span[@data-i18n = 'main:Log Out']").click()
        browser.quit()

        txns_df = pd.concat([txns_df, brokerage_interest_income])

        self.wb.sheets["Overview"].range( self.account1_name.replace(" ","_") ).value = float(account1_current_balance.replace("$","").replace(",","").strip())
        self.wb.sheets["Overview"].range( self.account2_name.replace(" ","_") ).value = float(account2_current_balance.replace("$","").replace(",","").strip())
        # Right here, you could write to a csv, instead of writing to the workbook
        self.wb.sheets["All Bank Transactions"].range('A1').options(pd.DataFrame, index = False).value = txns_df
        self.wb.sheets["All Bank Transactions"].range('A1').current_region.autofit()

    def refresh_income_and_expense_data(self): # change this to categories, or... income/expense generator

        df = self.wb.sheets["All Bank Transactions"].range("A1").current_region.options(pd.DataFrame).value
        df.reset_index(inplace = True)

        # Filter out all income expense excludes
        df = df[df["Income_Expense_Exclude"] == False]

        # Classify txns as either income or expense
        df.loc[(df["Account"] != self.credit_card_account_name) & (df["Credit_Debit_Ind"] == "Credit"), "Income_Expense_Ind"] = "Income"
        df.loc[(df["Account"] != self.credit_card_account_name) & (df["Credit_Debit_Ind"] == "Debit"), "Income_Expense_Ind"] = "Expense"
        df.loc[(df["Account"] == self.credit_card_account_name) & (df["Credit_Debit_Ind"] == "Credit"), "Income_Expense_Ind"] = "Expense"
        df.loc[(df["Account"] == self.credit_card_account_name) & (df["Credit_Debit_Ind"] == "Debit"), "Income_Expense_Ind"] = "Income"
        df.loc[(df["Account"] == "Robinhood Brokerage") & (df["Credit_Debit_Ind"] == "Credit"), "Income_Expense_Ind"] = "Income"

        # Flip the sign on all amounts to be positive (for credit card txns that show negetive amts)
        # Convert the column to a numeric type and format as Accounting (in Excel), too. - maybe do this later...
        df["Amount"] = df["Amount"].apply(abs)

        # drop these cols "Income_Expense_Exclude","Credit_Debit_Ind"
        df.drop(["Income_Expense_Exclude","Credit_Debit_Ind"], axis=1, inplace=True)

        # Rename the income/expense indicator col
        df.rename(columns={"Income_Expense_Ind":"Income or Expense"}, inplace=True)

        # Clean up the description col
        df["Description"] = df["Description"].apply(lambda x: remove_visa(x))
        df[['Txn Month/Day', 'Description']] = df.apply(
            lambda row: extract_and_remove_date(row['Description'], row['Date']),
            axis=1, result_type='expand'
        )

        # Add description category col
        df["Description_Category"] = ""
        df["Description_Category"] = df["Description"].apply(self.__categorize_description)
        # Add these description categories manually
        for index, row in self.manual_desc.iterrows():
            if ((df["Date"]==row[0]) & (df["Amount"]==row[1]) & (df["Description"]==row[2])).any():
                df.loc[ (df["Date"]==row[0]) & (df["Amount"]==row[1]) & (df["Description"]==row[2]), "Description_Category"] = row[3]

        # Rename the date col
        df.rename(columns={"Date":"Post Date","Description_Category":"Description Category"}, inplace=True)
        df = df[[
            "Post Date",
            "Txn Month/Day",
            "Account",
            "Amount",
            "Description",
            "Type",
            "Income or Expense",
            "Description Category"
        ]]
        
        # Replace txns for the HOA roof replacement
        incoming_txn = df[
            (df["Amount"] == 16815.39) & (df["Income or Expense"] == "Income") & (df["Post Date"] == "02/24/2025")
        ]
        outgoing_txn = df[
            (df["Amount"] == 17316.39) & (df["Income or Expense"] == "Expense") & (df["Post Date"] == "02/25/2025")
        ] 
        df.drop(incoming_txn.index, inplace = True)
        df.drop(outgoing_txn.index, inplace = True)
        new_txn = pd.DataFrame([{
            "Post Date": outgoing_txn["Post Date"].values[0],
            "Txn Month/Day": outgoing_txn["Txn Month/Day"].values[0],
            "Account": outgoing_txn["Account"].values[0],
            "Amount": 500,
            "Description": "Safeco Insurance Deductible - HOA Roof Replacement",
            "Type": outgoing_txn["Type"].values[0],
            "Income or Expense": "Expense",
            "Description Category": ""
        }])
        df = pd.concat([df, new_txn], ignore_index=True)

        # Sort the df by date (descending)
        df = df.sort_values(by = "Post Date", ascending = False)

        # Write the df to the Income and Expensess tab and make it a data table
        self.wb.sheets["Income and Expense Tracking"].tables("transactions").range.clear()
        self.wb.sheets["Income and Expense Tracking"].range('A1').options(pd.DataFrame, index = False).value = df
        self.wb.sheets["Income and Expense Tracking"].tables.add(source = self.wb.sheets["Income and Expense Tracking"].range("A1").current_region, name = "transactions")
        self.wb.sheets["Income and Expense Tracking"].range('A1').current_region.autofit()

    def get_investments_v1(self): 

        # provide option to pull all time investment data from Robinhood and Coinbase (from file...)

        # +++ Robinhood +++

        # Login
        login = rh.authentication.login(self.robinhood_u, self.robinhood_p) 

        # Get holdings data
        holdings_data = rh.account.build_holdings()
        df = pd.DataFrame(holdings_data)

        # Parse it out 
        df = df.transpose()
        df.reset_index(inplace = True)
        df.rename(columns = {"index":"symbol"}, inplace = True)
        df = df[["symbol","name","type","quantity","equity"]]
        df.rename(
            columns={
                "symbol":"Symbol",
                "name":"Name",
                "type":"Type",
                "quantity":"Quantity",
                "equity":"Current Equity"
            }, 
            inplace=True
        )

        # Log out
        rh.authentication.logout()

        # +++ Coinbase +++

        # Get all your crypto accounts
        client0 = Client("0", "0")
        client = RESTClient(api_key=self.coinbase_key_id, api_secret=self.coinbase_key_secret)
        crypto_accounts = client.get_accounts()["accounts"]
        # Build a list of tuples
        crypto_accounts_with_balances = []
        for crypto_account in crypto_accounts:

            if float(crypto_account["available_balance"]["value"]) > 0:
            
                crypto_symbol = crypto_account["available_balance"]["currency"]
                crypto_name = crypto_account["name"]
                crypto_quantity = crypto_account["available_balance"]["value"]
                crypto_exchange_rate = client0.get_exchange_rates(currency=crypto_symbol)["rates"]["USD"] # need to fix this
                crypto_equity = str(float(crypto_exchange_rate) * float(crypto_quantity))

                crypto_accounts_with_balances.append(
                    (
                        crypto_symbol,
                        crypto_name,
                        "cryptocurrency",
                        crypto_quantity,
                        crypto_equity  
                    )
                )

        df2 = pd.DataFrame(
            crypto_accounts_with_balances,
            columns = [
                "Symbol", "Name", "Type", "Quantity", "Current Equity"
            ]
        )

        # Pull out the USD amount
        usd_amt = df2[(df2["Symbol"]=="USD")].iloc[0]["Quantity"]
        df2.drop(index = df2[(df2["Symbol"]=="USD")].iloc[0].name, inplace = True)

        df = pd.concat([df,df2])

        # +++ Write it all to Excel +++

        # Write holdings data to the workbook and make it a table
        holdings_table_address = self.wb.sheets["Personal Investment Portfolio"].tables["holdings"].range.address
        self.wb.sheets["Personal Investment Portfolio"].range(re.sub(r'\$(\d+)$', increment, holdings_table_address)).delete(shift='up') 
        self.wb.sheets["Personal Investment Portfolio"].range("A1").options(index=False).value = df
        self.wb.sheets["Personal Investment Portfolio"].tables.add(source = self.wb.sheets["Personal Investment Portfolio"].range("A1").current_region, name = "holdings")
        self.wb.sheets["Personal Investment Portfolio"].range("A1").current_region.autofit()
        self.wb.sheets["Personal Investment Portfolio"].range("K6").value = usd_amt

    def retrieve_estatements(self):

        try:
                
            # Instantiate the webdriver object 
            chromeOptions = webdriver.ChromeOptions()
            settings = {
                "recentDestinations": [
                    {
                        "id": "Save as PDF",
                        "origin": "local",
                        "account": ""
                    }
                ],
                "selectedDestinationId": "Save as PDF",
                "version": 2
            }
            downloaded_estatement_folder = self.wb.sheets["Script Control Center & Ref Dta"].range("Downloaded_eStatement_folder").value
            prefs = {
                'printing.print_preview_sticky_settings.appState': json.dumps(settings),
                'savefile.default_directory': downloaded_estatement_folder
            }
            chromeOptions.add_experimental_option("prefs",prefs)
            chromeOptions.add_argument('--kiosk-printing')
            browser =  webdriver.Chrome(
                executable_path = self.wb.sheets["Script Control Center & Ref Dta"].range("Chromedriver").value, 
                options = chromeOptions
            )
            browser.implicitly_wait(10)

            # Login to OB (is there a way to use credentials that are saved in the browser???)
            browser.get('https://www.efirstbank.com/')
            browser.find_element_by_id('userId').send_keys(self.firstbank_u)
            browser.find_element_by_id('password').send_keys(self.firstbank_p)
            browser.find_element_by_id('logIn').click()

            # Define folder locations
            # -> root paths
            assets_and_liabilities = self.wb.sheets["Script Control Center & Ref Dta"].range("Assets_and_Liabilities_Path").value
            firstbank_asset_accounts = os.path.join(assets_and_liabilities, "Assets", "Bank Accounts", "FirstBank")
            firstbank_liability_account = os.path.join(assets_and_liabilities, "Liabilities", "FirstBank {account_name}".format(account_name = self.credit_card_account_name))
            # -> full directory paths (and a list of all those paths)
            account1_stmt_path = os.path.join(firstbank_asset_accounts, self.account1_name, "Current Statements in OB")
            account2_stmt_path = os.path.join(firstbank_asset_accounts, self.account2_name, "Current Statements in OB")
            account3_stmt_path = os.path.join(firstbank_asset_accounts, self.account3_name, "Current Statements in OB")
            credit_card_stmt_path = os.path.join(firstbank_liability_account,"Current Statements in OB")
            current_statement_in_ob_path_list = [account1_stmt_path,account2_stmt_path,account3_stmt_path,credit_card_stmt_path]

            # Navigate to the eStatements in online banking
            browser.find_element_by_xpath('//*[@id="obTab"]/a').click()
            browser.find_element_by_link_text('eStatements').click()

            xpath = '//*[@id="contentContainer"]/div[2]/div[2]/table/tbody/tr[{tr_index}]/td[{td_index}]' # /select
            for i in range(4):
                
                # Reference two siblings up from the parent to to the account name
                current_account = browser.find_element_by_xpath(xpath.format(tr_index = i+1, td_index = 1)).text
                print(current_account)
                
                current_account_dropdowns = browser.find_element_by_xpath(xpath.format(tr_index = i+1, td_index = 3)+"/select")
                date_options = current_account_dropdowns.find_elements_by_tag_name("option")
                for date_option in date_options:

                    # Select the statement date that you want to pull a statement for
                    statement_date = date_option.get_attribute("value")
                    statement_date = statement_date.replace("/","-")
                    date_option.click()
                    current_tab = browser.current_window_handle

                    # Click on the eStatement button for the estatement to show up in a new tab w/in the browser
                    browser.find_element_by_xpath(xpath.format(tr_index = i+1, td_index = 4) + '/div/input').click()

                    # switch into the new tab and wait for it to load
                    browser.switch_to.window(browser.window_handles[1])
                    embeded_web_element = browser.find_element_by_tag_name("embed")

                    # Print the page to pdf
                    browser.execute_script("window.print();")

                    # Folder reference will depend on...
                    if current_account == self.account1_name:
                        export_folder = account1_stmt_path
                    elif current_account == self.account2_name:
                        export_folder = account2_stmt_path
                    elif current_account == self.account3_name:
                        export_folder = account3_stmt_path
                    elif current_account == self.credit_card_account_name:
                        export_folder = credit_card_stmt_path

                    # Wait for estatementprep.do.pdf to be downloaded
                    time_threshold = 8
                    j = 1
                    # while not os.path.exists(os.path.join(downloaded_estatement_folder,"estatementprep.do.pdf")):
                    while not check_for_existing_pdf(downloaded_estatement_folder) and j < time_threshold:
                        time.sleep(2)
                        j += 1 
                    # Grab the name of the one file that should be in there
                    f_name = os.listdir(downloaded_estatement_folder)[0]
                    os.rename(
                        os.path.join(downloaded_estatement_folder,f_name),
                        os.path.join(export_folder,statement_date + ".pdf")
                    )
                    # wait for folder to be empty?
                    while check_for_existing_pdf(downloaded_estatement_folder) and j < time_threshold:
                        time.sleep(1)
                        j += 1 

                    browser.close()
                    browser.switch_to.window(current_tab)

                    #break # This is temporary

            # Log out and close both the browser and db cnxn
            browser.find_element_by_xpath("//span[@data-i18n = 'main:Log Out']").click()
            browser.quit()

            # +++ PDF merge routine +++
            
            # Loop through all folders holding bank statements
            for current_statement_in_ob_path in current_statement_in_ob_path_list:
                
                pdf_list = []
                for eStatement in os.listdir( current_statement_in_ob_path ):

                    pdf_list.append( os.path.join(current_statement_in_ob_path,eStatement) )

                # Merge all PDFs in the PDF list together
                eStatement_account = os.path.basename( os.path.split(current_statement_in_ob_path)[0] )
                PDFmerge(
                    pdf_list,
                    os.path.join(
                        os.path.abspath(os.path.join(current_statement_in_ob_path, os.pardir)),
                        'Merged {eStatement_account} eStatements.pdf'.format(eStatement_account = eStatement_account)
                    )
                )

            # write to log file
            with open(self.wb.sheets["Script Control Center & Ref Dta"].range("Log_File").value, 'w') as f:

                f.write("eStatements Retrieved Successfully")

        except Exception as e:

            # write to log file... 
            with open(self.wb.sheets["Script Control Center & Ref Dta"].range("Log_File").value, 'w') as f:
                f.write(str(e))
                f.write(traceback.format_exc())
