import os
import re
import sys
import fnmatch
import openpyxl
import ctypes
import time
from playwright.sync_api import sync_playwright
import getpass

# Define the show_elapsed_time function before main()
def show_elapsed_time(page, duration):
    start_time = time.time()
    elapsed_time = 0
    success = True
    last_log_index = 0
    short_timeout = 1000
    penter = 1

    while success and elapsed_time < duration:
        try:
            # Switch to default content and find the iframe
            frame = page.frame(name="main_frame")
            log_locator = frame.locator("#log")

            # Wait for the "okbutton" to be present
            fb2 = frame.locator("#okbutton")

            log_text = log_locator.text_content(timeout=short_timeout) or ""
            #print("Raw log_text"+log_text)

            log_length = len(log_text)
            #print("log_length"+ str(log_length))


            normalized_log_text = log_text.replace('\xa0', ' ')
            normalized_log_text = re.sub(r'\(\d{1,2}:\d{2}:\d{2} [APM]{2}  UTC \+\d{1,2}:\d{2}\)', r'\n\g<0>', normalized_log_text)

            log_entries = normalized_log_text.splitlines()

            #print(f"last_log_index: {last_log_index}, Total log entries: {len(log_entries)}")
            for i in range(last_log_index, len(log_entries)):
                entry = log_entries[i].strip()
                if entry:
                    if penter ==1:
                        print("", flush=True)
                        penter += 1
                    print(entry, flush=True)
                    #print(entry)
                    last_log_index = i + 1

            # Click the "okbutton" if found
            if fb2.is_visible(timeout=short_timeout):

                #Copying the above code befor button click to run it once more
                log_text = log_locator.text_content(timeout=short_timeout) or ""
                log_length = len(log_text)
                
                normalized_log_text = log_text.replace('\xa0', ' ')
                normalized_log_text = re.sub(r'\(\d{1,2}:\d{2}:\d{2} [APM]{2}  UTC \+\d{1,2}:\d{2}\)', r'\n\g<0>', normalized_log_text)

                log_entries = normalized_log_text.splitlines()
                for i in range(last_log_index, len(log_entries)):
                    entry = log_entries[i].strip()
                    if entry:
                        print(entry, flush=True)
                        last_log_index = i + 1

                fb2.click()  # This will click the button to finish the process
                success = False
                print("\nSuccessfully run the ruleset", flush=True)
                break
                
        except Exception as e:
            print(f"Error occurred: {str(e)}", flush=True)
            time.sleep(1)

        # Update elapsed time
        elapsed_time = time.time() - start_time
        minutes = int(elapsed_time // 60)
        seconds = int(elapsed_time % 60)
        elapsed_time_string = f"Waiting for server to be assigned, Elapsed time: {minutes:02d}:{seconds:02d}"
        if log_length == 0:
            print(f"\r{elapsed_time_string}", end="", flush=True)


    # Handle case if process takes longer than expected
    if elapsed_time >= duration:
        print("It is taking more than 15 mins, will be closing automatically, run task manually")
        ctypes.windll.user32.MessageBoxW(0, "It is taking more than 15 mins, will be closing automatically, run task manually", "Taking more time", 0)
        page.close()
        sys.exit()
    else:
        time.sleep(5)

def main():
    print("Please wait....")

    current_path = os.getcwd().replace('\\','/')

    # Locate the Excel file with 'Delivery_manager' in its name
    for filename in os.listdir(current_path):
        if fnmatch.fnmatch(filename, '*Delivery_manager*'):
            xlname = filename
    
    # Load the workbook and sheet
    wb = openpyxl.load_workbook(xlname, data_only=True)
    sh = wb["Parameters"]
    fcode = sh["B2"].value
    ncode = int(fcode)
    pcode = str(ncode)
    RulesetNum = sh["B6"].value
    Rulesetstr = str(RulesetNum)

    if not RulesetNum:
        print("Ruleset ID not found, enter the ruleset id in delivery manager and re-run")
        ctypes.windll.user32.MessageBoxW(0, "Ruleset ID not found, enter the ruleset id in delivery manager and re-run", "Ruleset ID not found", 0)
        return

    # Playwright browser automation
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        
        # Navigate to the page
        page.goto("https://author.euro.confirmit.com/confirm/authoring/Confirmit.aspx")

        # Login handling
        user_name_element = page.locator("#username")
        password_element = page.locator("#password")
        login_button = page.locator("#btnlogin")

        UserName = os.getlogin()

        # Set the username according to the current login
        if UserName == 'gregory':
            UserName1 = 'igoris.remeika'
        elif UserName == 'g.slavinskaite':
            UserName1 = 'gabriele.slavinskaite'
        elif UserName == 'tarun.kumar':
            UserName1 = 'tarun.kumar_sermo'
        else:
            UserName1 = UserName

        usernameauto = input(f"If your username is {UserName1}, enter y, else n (y/n): ")
        
        max_attempts = 5
        attempts = 0

        # Login loop
        while attempts < max_attempts:
            if usernameauto.lower() == "y":
                username = UserName1
            else:
                username = input("Enter your username: ")

            password = getpass.getpass("Enter your password: ")

            # Fill in the login details and submit
            user_name_element.fill(username)
            password_element.fill(password)
            login_button.click()

            # Wait for page load after login
            try:
                page.wait_for_selector("#__button_dataprocessing_inner", timeout=5000)
                print("Login successful!")
                break
            except:
                print("Invalid username or password. Please try again.")
                attempts += 1

        if attempts == max_attempts:
            print("Maximum login attempts reached. Exiting the program.")
            browser.close()
            return

        # After login, perform actions on the page as needed
        page.click("#__button_dataprocessing_inner")
        page.wait_for_selector("#__button_dp_rulesetlist", timeout=10000)
        page.click("#__button_dp_rulesetlist")
       
        # Switch to the main frame and wait for the Ruleset search field
        main_frame = page.frame_locator("#main_frame")
        main_frame.locator("body").wait_for(timeout=10000)
        
        # Wait for the Ruleset search field to appear, then input the Ruleset number
        search_field = main_frame.locator("#RulesetId")
        search_field.wait_for(timeout=10000)
        search_field.fill(Rulesetstr)
        search_field.press("Enter")

        #print("Ruleset ID entered")

        # Wait for the grid frame to load and locate it
        fgc_rule_set_frame = main_frame.frame_locator("#fgcRuleSet_gridframe")
        fgc_rule_set_frame.locator("body").wait_for(timeout=20000)

        # Try to find and click the ruleset link
        try:
            ruleset = fgc_rule_set_frame.locator("#fgRuleSet_theGrid_ctl02_aRuleSetName")
            ruleset.wait_for(timeout=10000)
            if ruleset.count() == 0:
                raise Exception(f"Ruleset ID {Rulesetstr} not found in SaaS.")
            button_text = ruleset.inner_text()
            ruleset.click()
        except Exception as e:
            print(f"Ruleset ID {Rulesetstr} not found in SaaS: {e}")
            ctypes.windll.user32.MessageBoxW(0, f"Ruleset ID {Rulesetstr} is not present in Confirmit SaaS", "Ruleset ID not found", 0)
            browser.close()
            sys.exit()
        
        # Execute the ruleset
        main_frame.locator("#ctl01_miExecuteItem").click()
        main_frame.locator("#okButton").click()
        print(f"Running ruleset {button_text}")

        # Wait for the server task to complete, or show elapsed time
        show_elapsed_time(page, 900)
        page.close()
        browser.close()



if __name__ == "__main__":
    main()




