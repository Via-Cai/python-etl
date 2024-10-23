import os
import time
import pandas as pd
from sqlalchemy import create_engine
import psycopg2

# Read Excel file
excel_file = r"C:\Users\llivi\OneDrive\Desktop\Employee_data.xlsx"
table_name = 'employee_data'

# Connect to PostgreSQL database
engine = create_engine('postgresql+psycopg2://postgres:admin@localhost/python-conn')

# Get the initial modified time
last_modified_time = os.path.getmtime(excel_file)

def update_postgres():
    try:
        df = pd.read_excel(excel_file)
        
        # Adding a FullName Column by combining first and last name
        df['FullName'] = df['FirstName'] + ' ' + df['LastName']

        # Write dataframe to PostgreSQL, if table already exist, replace it and index = False means it doesn't include the index of each row.
        df.to_sql(table_name, engine, if_exists = 'replace', index=False)
    except FileNotFoundError as e:
        print(f"Error: Excel file not found. {e}")
    except psycopg2.Error as e:
        print(f"Error while connecting to PostgreSQL or executing SQL. {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
    else:
        print(f"Data written to the PostgreSQL table '{table_name}' successfully!")
    finally:
        print("Execution finished.")

while True:
    try:
        current_modified_time = os.path.getmtime(excel_file)
        
        if current_modified_time != last_modified_time:
            print("Excel file has been modified. Updating PostgreSQL...")
            update_postgres()
            last_modified_time = current_modified_time # updete the modified time
            
        time.sleep(15)
    
    except KeyboardInterrupt:
        print("Stopped Monitoring.")
        break