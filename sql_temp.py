'''
World Temperature Database Cities Query Script
Author: Hashim-Jones, Jake (21/09/2017)
Version 1.0.

This script creates a connection to the database created by db_create.py. It then queries the database for a list of all cities in the Southern Hemisphere, and adds this to a new database table. It also prints the minimum, maximum and average temperature data for the state of Queensland (Australia) in the year 2000.

See readme for more details.
'''

########################################
## Importing and Creating Definitions ##
########################################

print("Initializing.\n")
import sqlite3
from os.path import isfile
import datetime

def yesNoInput(prompt=""):
    while True:
        answer = input(prompt).upper() # Get and capitalize user input.
        if answer == "Y" or answer == "N":
            break
        print("Error! Please enter 'Y' or 'N' (case-insensitive)") # Reprompt otherwise.
    return answer == "Y" # Return True if the user input 'Y' (meaning yes), or False if the user input 'N'.

################################
## Create Database Connection ##
################################

if not isfile("Temperature_Data.db"): # Check that database file exists...
    print("Error, 'Temperature_Data.db' does not exist. Please run 'db_create.py' first.") # Prompt to run the creation script if it does not.
    exit(0)

dbConnection = sqlite3.connect("Temperature_Data.db") # Open connection to the (existing) database.
print("\nOpening 'Temperature_Data.db'")
print("Connected to database.", datetime.datetime.now(),"\nSqlite version:", sqlite3.sqlite_version, "\n")
dbCursor = dbConnection.cursor() # Create cursor object.

print("Checking Tables...\n")
existingTables = dbCursor.execute("Select Name From sqlite_master Where type = 'table';").fetchall() # Obtaining table names that exist in the database.
existingTables = [name[0] for name in existingTables] # Extracting table names from the database output.
missingTables = [name for name in ['Country', 'MajorCity', 'State'] if name not in existingTables] # Compile list of 'expected' tables in the database that are NOT present.

if len(missingTables) >0: # This branch is executed if the database is mising one of the essential 3 tables.
    for name in missingTables:
        print("'{}' Table is missing.".format(name)) # Alert user that table(s) are missing.
    print("\nError, 'Temperature_Data.db has incomplete data. Please run 'db_create.py' to ensure all appropriate data is available for the program.\n") # Prompt user to run the creation script.
    print("Disconnecting from the database...")
    dbConnection.close() # Close the database connection.
    print("Disconnected from the database.", datetime.datetime.now())
    exit(0) # Terminate the program.
else: # Everything is fine and the program continues.
    print("Success.Check Complete. Database has all appropriate data.\n\n")

########################
## Query the Database ##
########################

print("Generating a list of distinctive major cities in the Southern Hemisphere...\n")
city = dbCursor.execute('''
SELECT distinct City, Country, Latitude, Longitude 
FROM MajorCity 
WHERE Latitude LIKE '%S'
ORDER BY Country;''').fetchall()

for data in city: #Print all data from the database that will be entered into the new table.
    print("     {}, {} ({} {})".format(data[0], data[1], data[2], data[3]))

print("\nAdding this data to new table.\n")
newTableSchema='''
Create Table "Southern Cities"(
	city varchar2(30),
	country varchar2(20),
	latitude varchar2(9) NOT NULL,
	longitude varchar2(9) NOT NULL,
	CONSTRAINT southerncities_PK PRIMARY KEY(city, country)
);''' #Definition of new table.

if 'Southern Cities' in existingTables: # Check that the new table does not exist.
    if not yesNoInput("'Southern Cities' Table already exists. Continuing will override the table and all of its data. Would you like to continue (Y/N)?"): # Give user the option to quit.
        print("\nDisconnecting from the database...")
        dbConnection.close()
        print("Disconnected from the database.", datetime.datetime.now())
        exit(0)
    else:
        print("\n'Southern Cities' will be overridden.\n")
        print("Dropping Table - 'Southern Cities'")
        dbCursor.execute('Drop Table "Southern Cities";') # Remove table from database.
        print("Table Dropped!\n")

dbCursor.execute(newTableSchema) # Create new table.
print("'Southern Cities' Table has been created!\n")

insertRecords='''
Insert Into "Southern Cities" Values ('{city}','{country}','{latitude}','{longitude}'); '''
print("Exporting Data to 'Southern Cities' Table...\n")
for city in city:
    dbCursor.execute(insertRecords.format(city=city[0], country=city[1], latitude=city[2], longitude=city[3])) # Add each record to the new table.
    print("     Succceess. Record added!")

query='''
SELECT min(AverageTemperature), max(AverageTemperature), avg(AverageTemperature) 
FROM State 
WHERE state='Queensland' AND
    country='Australia' AND
	Date BETWEEN '2000-01-01' AND '2000-12-31';
''' # Query for statistics from the database.

print("\nRetrieving statistical data for average temperatures in 'Queensland, Australia' in the year 2000...")
queenslandStats=dbCursor.execute(query).fetchone()
print('''
    Minimum Temperature: {min}
    Maximum Temperature: {max}
    Average Temperature: {avg}
'''.format(min='{0:.3f}'.format(queenslandStats[0]), max='{0:.3f}'.format(queenslandStats[1]), avg='{0:.3f}'.format(queenslandStats[2]))) # Print the retrieved statistics.


##################################################
## Commit Database Changes and Close Connection ##
##################################################
print("\nCommiting changes to the database...")
dbConnection.commit()
print("Success.\n")
print("Disconnecting from the database...")
dbConnection.close()
print("Disconnected from the database.",datetime.datetime.now())
