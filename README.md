This is a useful tool for people with lengthy and numerous emails who cannot just be searching each email one by one.

This python project uses Microsoft Graph API to help you filter your account(s) for emails that contain certain search terms in their subject, content, from or to address, by just using your username, password, client\_id, and client\_secret. It automates the entire login process into Azure, gets the authorization\_code, generates a token, and connects to your mail folders to perform the filtering and then saves matching emails to RESULTS folder, telling you which folder they were found.

Everything is already set up; you don't need to install anything.

However:

1. You should have already done an app registration from Microsoft Azure, granted the app permissions needed (User.Read and Mail.Read) and gotten the client\_id and client\_secret in the process. check for steps [here (https://helpdesk.kaseya.com/hc/en-gb/articles/11471626278545-How-to-set-up-an-app-registration-for-Microsoft-Graph-synchronization-only)
   ](https://helpdesk.kaseya.com/hc/en-gb/articles/11471626278545-How-to-set-up-an-app-registration-for-Microsoft-Graph-synchronization-only)
2. Make sure you have python installed and the default python version of your system is python3.9.x (x can be anything).
   NB: If you have some technical knowledge, you will quickly find out that your choice of options are not limited to the ones I provided; you can use any python version, or even use in a virtual environment (which I did when developing this software). But note that I used python 3.9 in a development environment.
3. In the same folder you see this script, run "pip install django" and "python manage.py runserver"
4. You should see some messages including this "Starting development server at http://127.0.0.1:8000/
   Quit the server with CONTROL-C."
5. your Microsoft Graph API redirect_url is "http://127.0.0.1:8000/code" or "http://localhost:8000/code"; It's same thing.
6. Then (from command line/powershell/vscode terminal) move to the "FILTER" folder (cd FILTER) inside the first "FILTER" folder and run "python engine.py"
   NOTE: You might need to run "python3 engine.py" instead; especially since you are using a Mac.
7. Before then make sure your variables are already defined:
   a. specify your "username(separator)password(separator)client_id(separator)client_secret in input.txt
   b. specify your separator in "separator.txt". The default is ","
   c. leave the "All" in folders.txt if you want to search all folders including junk. Else, leave folders.txt empty
   d. If any matches are found in your email account, the email account will be saved to "found.txt"
   e. Search the email "to address" by specifying your search terms in separate lines of "search_to.txt"
   f. Search the email "from address" by specifying your search terms in separate lines of "search_from.txt"
   g. Search the email "body" by specifying your search terms in separate lines of "search_content.txt"
   h. Search the email "subject" by specifying your search terms in separate lines of "search_subject.txt"
8. Valid emails will be saved as .eml files, with a hint on the account, and mail folder they were found. It's a hint, but you can disable this action by changing the "1" in 'choice.txt' to  "0".
9. You can filter emails from multiple accounts by specifying each account's username(separator)password(separator)client_id(separator)client_secret in separate lines of "input.txt".
# graphAPI-email-filter-Mac-


## UPDATE
Updated chromedriver to 119.