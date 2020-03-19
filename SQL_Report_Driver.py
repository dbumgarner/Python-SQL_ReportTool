"""Driver program that executes sets of queries provided in a subsequent file based on user controls"""
__author__      = "David Gray"
__status__      = "BETA"
__version__     = "2.6"
__maintainer__  = "David Gray"
__email__       = "DGray@GraylanderLabs.com"
__phone__       = "512.694.5678"

# Imports
import pyodbc
import csv
import sys
import getpass
import colorama
import decimal
from datetime import datetime
import os
import errno
from colorama import Fore
colorama.init(autoreset=True)
import report_queries

# Global variables
server = 'ServerLocationGoesHere' 
database = 'DatabaseNameGoesHere' 
os.system('mode con: cols=160')
user_filtering = True
original_asignee = True
query = ''

# Helper function providing exit procedure for pyinstaller so it doesn't close immediately on exit/quit
def graceful_exit(s):
    print(Fore.RED + s)
    quit()
    input( " \nPress enter to continue...")
    quit()

# Provide formatting for warnings 
def warn(s, border):
    if border:
        border = ''
        for i in range(0,len(s)+6):
            border+='+'
        print(Fore.RED + border)
        print(Fore.RED + '   ' + s)
        print(Fore.RED + border)
    else:
        print(Fore.RED + s)

# Provide formatting for alerts  
def alert(s, border):
    if border:
        border = ''
        for i in range(0,len(s)+6):
            border+='+'
        print(Fore.YELLOW + border)
        print(Fore.YELLOW + '   ' + s)
        print(Fore.YELLOW + border)
    else:
        print(Fore.YELLOW + s)
        
# Provide formatting for Notifications 
def notify(s, border):
    if border:
        border = ''
        for i in range(0,len(s)+6):
            border+='+'
        print(Fore.GREEN + border)
        print(Fore.GREEN + '   ' + s)
        print(Fore.GREEN + border)
    else:
        print(Fore.GREEN + s)
        
# Function that sets user filtering equal to the input
def set_user_filtering(uf):
    global user_filtering 
    user_filtering = uf
    reporting()
    
# Function that toggles the filter target    
def toggle_user_target():
    global original_asignee
    original_asignee = not original_asignee
    reporting()

# Helper function that returns the ID of a username selection to ensure it exists and to be used in a where clause   
def lookup_user_id(uname):
    cursor.execute("select id from users where username like '"+str(uname)+"'")
    return cursor.fetchall()[0][0]
    
# Function that executes the selected query and returns the result list
def execute_query(choice):
    global user_filtering
    if user_filtering == False: # If filtering is disabled execute the defualt query
        cursor.execute("set nocount on; "+report_queries.queries[choice][1])
    elif original_asignee: # If marked as the original asignee then replace the wildcard with the username
            cursor.execute("set nocount on; "+report_queries.queries[choice][1].replace("username like '%'"," username like '"+user_filtering+"'"))
    else: # Otherwise it's against the updated user so replace the wildcard with the ID of the username
            cursor.execute("set nocount on; "+report_queries.queries[choice][1].replace("username like '%'"," updated_user_id like '"+str(lookup_user_id(user_filtering))+"'")) 
    return cursor.fetchall()
    
# Write function that outputs CSVs if there is need for it    
def write_file(rows, choice):
    global warning_tripped    
    warning_tripped = False
    
    if len(rows)>0:    # Only write if there is data to write to the file
        path_file= str(datetime.today().strftime('%Y-%m-%d')+"\\"+report_queries.queries[choice][0]+".csv")     # Establish the file that is going to be written based on date and report name
        if not os.path.exists(os.path.dirname(path_file)):                                                      # If it doesn't exist   
            try:                                                                                                    # Try to make it
                os.makedirs(os.path.dirname(path_file))
            except OSError as exc:                                                                                  # Guard against race condition 
                if exc.errno != errno.EEXIST:
                    raise
        try:            
            with open(path_file, 'w', newline='') as outfile:                                                       # Open the new file
                writer = csv.writer(outfile)                                            
                writer.writerow(report_queries.queries[choice][2])                                                  # Write the headers
                for row in rows:                                                                                        # Write each of the rows
                    writer.writerow(row)                                                                            
                outfile.close()                                                                                     # Close the file
        except:
            warning_tripped = True
            warn("WARNING: Could not write report! Do you have it open in another program?", True)

# Print function for if the user wants to see the output        
def print_rows(rows, choice):
    record_count = 1
    
    # For each row display the header and the relevant data horizantally 
    for row in rows:
                print(Fore.CYAN + "Record "+str(record_count))
                record_count+=1
                for i in range(0, len(report_queries.queries[choice][2])):
                    print('{:<20s}{:<20s}'.format(str(report_queries.queries[choice][2][i]),str(row[i])))
                print("-------------------------------------------------------") 
    if len(rows) == 0:
        notify("No records in report!", True)
    else: 
        print(Fore.GREEN + "Total records: " + str(len(rows))+"\n")

# Pull all analyst reports based on filtering    
def all_analyst_reports():
    global user_filtering
    print()
        
    for i in range(0, len(report_queries.queries)):           # For all of the reports,
       if report_queries.queries[i][3] == "Case Analyst":       # If it's a Case Analyst report
           print("  Running report: "+ str(i))                      # Notify the report is running
           write_file(execute_query(i), i)                          # Write a file of the result of that query (assuming there is data)
   
    # If no files were written alert that no files need attention
    if not os.path.exists(datetime.today().strftime('%Y-%m-%d')):
        notify('No files need your attention today!', True)
    # Otherwise alert the user of the files locations  
    
    else:
        notify("Files can be located in the " + str(datetime.today().strftime('%Y-%m-%d') + " subdirectory contained in the directory with this tool."), True)
    # Return to main menu
    reporting()
    
# Main body of the tool that loops a repeating menu prompting which report should be pulled or settings should be adjusted   
def reporting():    
    global user_filtering
    choice = ''
    # Loop until user exit
    while True:
        # Prompt dynamically for report selection from reports provided in report_queries.py
        print(Fore.YELLOW + "\nWhich report would you like to run?")
        for option in range(0, len(report_queries.queries)):
            print("  "+str(option)+":\t"+report_queries.queries[option][0])
        
        # Display what the current report filter is so there's no ambiguity as to what data will be pulled
        if user_filtering==True:
            print(Fore.RED+ "\n  Current Filter:  "+username)
        elif user_filtering==False:
            print(Fore.RED+ "\n  ALL FILTERING DISABLED")   
        else:
            print(Fore.RED+ "\n  Current Filter:  "+ str(user_filtering))
            
        # Display current original asignee
        if user_filtering!=False:
            if original_asignee:
                print(Fore.RED+ "  Filter Target:   Original Asignee")
            else:
                print(Fore.RED+ "  Filter Target:   Most Recent Updater")
        
        # Offer running all reports in a batch, and adjusting the filter
        print(Fore.CYAN+"\n  A/a:\tRun All Case Analyst Reports Fitting Filter")
        print(Fore.CYAN+"  D/d:\tDisable Filtering")
        print(Fore.CYAN+"  R/r:\tReset Filtering to Self (DEFAULT)")
        print(Fore.CYAN+"  S/s:\tSet Filtering to Username")
        print(Fore.CYAN+"  T/t:\tToggle filtering against Original Asignee/Most Recent Updater")
        print(Fore.CYAN+"  Q/q:\tExit Tool")

        # Try to grab report selection from user as an integer to target report by index
        # Exceptions are where submenu functions are handled as listed below
        choice=input('\n -> ')
        try:    
            choice = int(choice)
            if choice == -1:
                graceful_exit("")
            if not 0 <= choice < len(report_queries.queries):
                raise
            
        except:
            if choice in ['A','a']:             # Run all analyst reports for the current filter
                all_analyst_reports()               
            elif choice in ['D','d']:           # Disable user_filtering
                set_user_filtering(False)            
            elif choice in ['R','r']:           # Reset user_filtering to the logged in user
                set_user_filtering(username)                
            elif choice in ['S','s']:           # Set user_filtering to a target username
                print(Fore.YELLOW + "\nEnter the username you want to filter against.")
                user = input(" -> ")
                try:
                    if lookup_user_id(user) is not None:
                        set_user_filtering(user)
                except: 
                    print("That is not a valid user")
                    reporting()            
            elif choice in ['T','t']:           # Toggle against updated_user_id
                toggle_user_target()
            elif choice in ['Q','q']:           # Exit
                graceful_exit("Exiting Tool")
            else:                               # Any other input defaults to looping the menu again
                print(Fore.RED+"\nBad report selection.")
                reporting()
         
        #Execute the query if a single report was chosen
        rows = execute_query(choice)

        # Loop through collected data and format it to the console and show an entry count unless there are >10 rows, in which case prompt to display
        print("\n")
        if len(rows)>10:
        
            
            warn("WARNING: There are "+ str(len(rows)) + " rows.", True)
            print(Fore.YELLOW + "Do you want to display all of them?")
            opt = input("Yy/Nn: ")
            if opt in ['Y','y']:
                print_rows(rows, choice)
        else: 
            print_rows(rows, choice)
        
        # Offer to export to Excel if there are returned rows and grab choice. If in Y,y then write to a CSV
        if len(rows)>0:
            sys.stdout.write(Fore.YELLOW+"Export to Excel? Yy/Nn:  ")
            opt = input()
            if opt in ['Y','y']:
               write_file(rows, choice)
               if not warning_tripped:
                    print(Fore.GREEN + "\nFile is entitled \""+report_queries.queries[choice][0]+".csv\" \nand can be located in the " + str(datetime.today().strftime('%Y-%m-%d') + " subdirectory contained in the directory with this tool."))

##### BEGIN PROGRAM CONTROL FLOW #####
logged_in=False

print(Fore.GREEN
      +'Version:\t\t'+__status__+' v'+__version__
      +'\nSupport & Maintenance:  Contact '+__maintainer__+' at '+__email__+' or '+__phone__+'\n')
      
while not logged_in:
    # Pull username and quiet password from user
    print(Fore.YELLOW+"\nEnter your SQL Credentials. The password will appear blank while typing for security purposes.\n")
    username = input("\tUsername: ")
    password = getpass.getpass(prompt="\tPassword: ")
    user_filtering = username
    # Try to set up connection and catch if it fails
    try:
        cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
        cursor = cnxn.cursor()
        print(Fore.GREEN+"\n\n   :::Credentials Accepted:::\n") 
        logged_in=True
    except Exception as e:
        print(Fore.RED+"\n\n   :::Bad credentials or database connection:::")
        logged_in=False
 
# Enter main loop REPORTING
reporting()
    
 