from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from msal import ConfidentialClientApplication
import os
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
import requests
import time
import re

# Define the scopes that you need to access
SCOPES = ["User.Read", "Mail.Read"]

# Read separator character from file
with open('separator.txt', 'r') as f:
    separator = f.read().strip()

# choose whether to save the emails and their location directories in RESULTS folder.
with open('choice.txt', 'r') as f:
    choice = f.read().strip()

#Read choice of mail folders from file
with open('folders.txt', 'r') as c:
    choose = c.read().strip()

# Read search terms from files
search_subjects = [term.strip() for term in open('search_subject.txt').readlines() if term.strip()]
search_contents = [term.strip() for term in open('search_content.txt').readlines() if term.strip()]
search_froms = [term.strip() for term in open('search_from.txt').readlines() if term.strip()]
search_tos = [term.strip() for term in open('search_to.txt').readlines() if term.strip()]

# Create RESULTS directory if it doesn't exist
if not os.path.exists('RESULTS'):
    os.makedirs('RESULTS')

#Main engine begins here
with open('input.txt', 'r') as input_file:
    for line in input_file:
        # Split line into all input portions
        username, password, client_id, client_secret= map(str.strip, line.split(separator))
        print("the following credentials below are provided for you to verify...")
        print(username)
        print(password)
        print(f"client_id : {client_id}")
        print(f"client_secret : {client_secret}")

         # Set the path to the ChromeDriver executable
        chrome_driver_path = 'chromedriver'

        # Configure Chrome options
        chrome_options = Options()

        # Create a new instance of the ChromeDriver service
        service = Service(chrome_driver_path)

        # Create a new instance of the Chrome driver
        driver = webdriver.Chrome(service=service, options=chrome_options)

        driver.maximize_window()
        driver.implicitly_wait(10)

        try:
            # Clear cookies
            driver.delete_all_cookies()
            
            #Visit the main Azure portal
            driver.get('https://portal.azure.com')
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, 'loginfmt')))

            # Fill in email
            email_input = driver.find_element(By.NAME, 'loginfmt')
            email_input.clear()
            email_input.send_keys(username)
            email_input.send_keys(Keys.ENTER)

            # Wait for the "Work or school account or personal account" page to load and choose "Work or school account"
            try:
                WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'aadTile')))
                driver.find_element(By.ID, 'aadTile').click()
            except:
                pass

            # Wait for password page to load
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, 'passwd')))

            #Clear the password (if not empty) and paste the real password
            password_input = driver.find_element(By.CSS_SELECTOR, '#i0118')
            password_input.clear()
            password_input.send_keys(password.strip())
            password_input.send_keys(Keys.ENTER)
            
            # If the "stay signed in?" page pops up, choose "No"
            stay_signed_in_prompt = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#idSIButton9')))
            if stay_signed_in_prompt.is_displayed():
                stay_signed_in_prompt.click()
                WebDriverWait(driver, 10).until(EC.invisibility_of_element_located((By.CSS_SELECTOR, '#idSIButton9')))
            else:
                continue

            # Wait for the portal to load, in order to ensure the login is completed
            try:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'mainContainer')))
            except:
                print("page not loading")

            # Get the current window handle
            main_window_handle = driver.current_window_handle

            # Keep attempting to open a new window and tab until they are created
            driver.execute_script("window.open('');")

            # Switch to the new window and load a link
            driver.switch_to.window(driver.window_handles[1])

            #Initialise the Graph API connection point
            client = ConfidentialClientApplication(client_id=client_id, client_credential=client_secret)
            authorization_url = client.get_authorization_request_url(SCOPES)
            driver.get(authorization_url)

            # This very important function is responsible for handling the generation of the authorization code.
            def generate_auth_code():
                try:
                    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'appConfirmContinue')))
                    driver.find_element(By.ID, 'appConfirmContinue').click()
                except:
                    pass

                try:
                    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'idBtn_Accept')))
                    driver.find_element(By.ID, 'idBtn_Accept').click()
                except:
                    pass

                # Wait for the authorization code to be in the URL.
                WebDriverWait(driver, 30).until(EC.url_contains("code="))

                # Get the authorization code from the URL.
                url = driver.current_url
                match = re.search(r'code=(.+)', url)
                auth_code = match.group(1)
                return auth_code
            
            # Store the returned auth_code as a variable for later use
            auth_code = generate_auth_code()
            print(auth_code)

            # Generate an access token from the auth_code
            access_token = client.acquire_token_by_authorization_code(code=auth_code, scopes=SCOPES)
            access_token = access_token['access_token']
            
            # Create the request body.
            headers = {
                'Authorization': 'Bearer ' + access_token,
                'Content-Type': 'application/json'
            }

            # Fetch messages from all mail folders
            response = requests.get('https://graph.microsoft.com/v1.0/me/mailFolders', headers=headers)
            mail_folders = response.json()['value']
            print(mail_folders)

            # Determine which folders to search
            if choose == "All":
                folders_to_search = mail_folders
            else:
                folders_to_search = [folder for folder in mail_folders if folder['displayName'] != 'junk email' and folder['displayName'] != 'Junk Email']
            
            # Search selected folders
            found_emails = []
            for folder in folders_to_search:
                folder_name = folder['displayName']
                folder_id = folder['id']
                messages_url = f'https://graph.microsoft.com/v1.0/me/mailFolders/{folder_id}/messages'
                messages_response = requests.get(messages_url, headers=headers)
                messages = messages_response.json()['value']
                for message in messages:
                    # Check if message matches search criteria
                    if any(term in message['body']['content'] for term in search_contents) or \
                            any(term in message['from']['emailAddress']['address'] for term in search_froms) or \
                            any(any(term in recipient['emailAddress']['address'] for recipient in message['toRecipients']) for term in search_tos) or \
                            any(term in message['subject'] for term in search_subjects):
                        found_emails.append(message)
                        print(f"match found in {folder_name}...")

                        # Save email as .eml file
                        email_id = message['id']
                        email_subject = message['subject']
                        base_url = 'https://graph.microsoft.com/v1.0'
                        email_url = f'{base_url}/me/messages/{email_id}/$value'
                        email_response = requests.get(email_url, headers=headers)
                        email_content = email_response.text
                        if choice == "1":
                            try:
                                username_dir = os.path.join('RESULTS', username)
                                if not os.path.exists(username_dir):
                                    os.makedirs(username_dir)
                                folder_dir = os.path.join(username_dir, folder_name)
                                if not os.path.exists(folder_dir):
                                    os.makedirs(folder_dir)
                                email_path = os.path.join(folder_dir, f'{email_subject}.eml')
                                with open(email_path, 'w') as f:
                                    f.write(email_content)
                            except:
                                pass
                        else:
                            print("Email results will not be saved, and you might not remember which folders they were found.")
                            print("To enable saving, change 'choice.txt' value to 1")

            # Save username and password to found.txt or not_found.txt depending on whether any emails matches were found
            if found_emails:
                with open('found.txt', 'a') as f:
                    f.write(f'{username}{separator}{password}\n')
            else:
                with open('not_found.txt', 'a') as f:
                    f.write(f'{username}{separator}{password}\n')
        except Exception as e:
            print(f"Error occurred for line: {line.strip()}")
            print(f"Error message: {str(e)}")
            continue
        finally:
            print("Moving to the next account...")

            # Quit the driver after every account access to further prevent any cookies from interfering with a new login
            driver.quit()

# Finally, the driver is quit to ensure that the code ends properly.
driver.quit()