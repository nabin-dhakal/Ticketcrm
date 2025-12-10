#TicketCRM

This is the dynamic ticket generation and mailing script.

It template image file and information from the excel sheet and generate a ticket to unique user and mail them using gmail service.


Procedure to execute this script:
-    1.) Create a .env file with reference from .env.example file ( requirements: generate a app password on the gmail used for sending mail.)
-    2.) Create a Virtual environment in python.
-    Command to execute :
-           python -m venv .venv

    In linux and Mac:
-        source .venv/bin/activate

-    In windows:
-        .\venv\scripts\activate.bat

-    3.) Install the required library from the requirements.txt file.


-    Command to execute:
-        pip install -r requirements.txt
    
-    4.) Import excel file, template image and also create a output folder. 

-    5.) Execute the script:

-   command:
-            python script.py
