# excel_agent_step_A.py

# --- Step 1: Import all the necessary libraries ---
import os
import gspread
from gspread_dataframe import get_as_dataframe
from google.oauth2.service_account import Credentials
from langchain_google_genai import ChatGoogleGenerativeAI
from langchain_experimental.agents.agent_toolkits import create_pandas_dataframe_agent
from dotenv import load_dotenv  # <-- ADD THIS LINE

load_dotenv()  # <-- ADD THIS LINE TO LOAD VARIABLES FROM .env



# --- Step 2: Configuration - SET YOUR DETAILS HERE ---

# The name of the Google Sheet file you want to access.
# Example: "My Company Sales Data"
GOOGLE_SHEET_KEY = '13eZQCeGZwDJRk8-yQuIQMCtroZ62BCDmSMS7rClkhIU'  # <--- CHANGED

# The name of the specific worksheet (tab) within that file.
# Example: "Q3_2024" or "Sheet1"
WORKSHEET_NAME = 'Creating_Tables'            # <--- CHANGE THIS to your worksheet's name

# The path to your Google Cloud service account JSON file.
# This file allows your script to securely access your Google Sheet.
# See the setup instructions below on how to get this file.
SERVICE_ACCOUNT_FILE = 'credentials.json' # <--- Make sure this file is in the same directory

# --- Step 3: Function to connect and fetch data from Google Sheets ---

def get_data_from_google_sheet(sheet_key, worksheet_name):
    """
    Connects to Google Sheets using a service account and loads a specific
    worksheet into a pandas DataFrame.
    """
    try:
        print("Authenticating with Google Sheets...")
        # Define the scope of access needed for the APIs
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        # Authenticate using the service account credentials
        creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=scopes)
        gc = gspread.authorize(creds)

        print(f"Opening Google Sheet: '{sheet_key}'...")
        # Open the spreadsheet by its name
        spreadsheet = gc.open_by_key(sheet_key) # <--- CHANGED

        
        print(f"Accessing worksheet: '{worksheet_name}'...")
        # Select the specific worksheet by its name
        worksheet = spreadsheet.worksheet(worksheet_name)
        
        print("Fetching data and loading into DataFrame...")
        # Convert the worksheet data into a pandas DataFrame
        df = get_as_dataframe(worksheet)
        
        # A common issue is gspread reading empty rows. Let's drop them.
        df.dropna(how='all', inplace=True)
        
        print("Data loaded successfully!")
        return df

    except FileNotFoundError:
        print(f"ERROR: The service account file '{SERVICE_ACCOUNT_FILE}' was not found.")
        print("Please ensure the file is in the same directory and the filename is correct.")
        return None
    except gspread.exceptions.SpreadsheetNotFound:
        print(f"ERROR: The Google Sheet '{sheet_key}' was not found.")
        print("Please check the name and make sure you have shared it with the service account email.")
        return None
    except gspread.exceptions.WorksheetNotFound:
        print(f"ERROR: The worksheet '{worksheet_name}' was not found in the spreadsheet.")
        return None
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        return None

# --- Step 4: Main execution block ---

if __name__ == "__main__":
    print("--- Starting Step A: The 'Baby' Google Sheets Agent ---")

    # Make sure you have set your Google AI API key as an environment variable
    # Example: export GOOGLE_API_KEY="your_api_key_here"
    if 'GOOGLE_API_KEY' not in os.environ:
        print("ERROR: GOOGLE_API_KEY environment variable not set.")
        print("Please set your API key to run the agent.")
    else:
        # Load the data from the Google Sheet into our DataFrame
        df = get_data_from_google_sheet(GOOGLE_SHEET_KEY, WORKSHEET_NAME)

        # Only proceed if the DataFrame was loaded successfully
        if df is not None:
            # print("\n--- Data Preview ---")
            # print(df.head())
            # print("\n--------------------\n")

            # Initialize the LLM "brain" of our agent
            llm = ChatGoogleGenerativeAI(model="gemini-2.5-flash", temperature=0)

            # Create the special pandas agent. This bundles the LLM with the DataFrame
            # and gives it the ability to generate and run python code on the data.
            # `verbose=True` is critical for learning! It shows the agent's thought process.
            agent = create_pandas_dataframe_agent(
            llm,
            df,
            agent_type="zero-shot-react-description",
            verbose=False,
            allow_dangerous_code=True
        )

        # --- NEW --- Let's try two different prompts to see how formatting changes
        
        # --- List of 10 Test Prompts ---
        test_prompts = [
            # 1. Simple Filter
            "Show me all employees who work in the HR department.",
            # 2. Compound Filter (Logic & Date)
            "Find employees in the 'AC' department who were hired before the year 2000.",
            # 3. Sorting
            "List all employees, sorted by their hire date from the most recent to the oldest.",
            # 4. Simple Aggregation (Count)
            "How many employees work in Building 1?",
            # 5. Grouped Aggregation (Very Common Task)
            "What is the total number of employees in each department? List the department and the count.",
            # 6. Search with a String Condition
            "Find all employees whose last name starts with 'S'.",
            # 7. Multi-step Reasoning (Implicit Logic)
            "Who is the most recently hired employee in the entire company?",
            # 8. Ambiguous Column Name (Synonym Test)
            "Sort the entire list by employee identification number in descending order.",
            # 9. Update an Existing Row (Modification Test)
            "Find the employee with Emp ID 1075 and change their Location to 'Building 4'.",
            # 10. Add a New Row (Addition Test)
            "Add a new employee: Emp ID 2001, Last Name 'Turing', First Name 'Alan', Dept 'AC', E-mail 'alant', Phone Ext 100, Location 'Building 2', Hire Date '5/15/2024'."
        ]

        # --- Execute the Test Suite ---
        for i, prompt in enumerate(test_prompts):
            print(f"\n===================[ Test Prompt #{i+1} ]===================")
            print(f"User Prompt: {prompt}")
            print("---------------------------------------------------------")
            
            try:
                response = agent.invoke({"input": prompt})
                print("\n--- Agent's Final Answer ---")
                print(response['output'])
            except Exception as e:
                print(f"\n--- An Error Occurred ---")
                print(f"The agent failed on this prompt. Error: {e}")
            
            print("=========================================================\n")