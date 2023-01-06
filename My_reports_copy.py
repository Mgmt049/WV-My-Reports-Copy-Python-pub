import pandas as pd
from ldap3 import Server, Connection, ALL #, ObjectDef, Reader, Tls, SAFE_SYNC
from ldap3.core.exceptions import LDAPCommunicationError, LDAPBindError#, LDAPException #NOTE: The modules in ldap3 are organized into packages & must use this dot hierarchy
import tkinter as tk
from tkinter.filedialog import askopenfilenames, askopenfilename
import tkinter.filedialog
import numpy as np
import shutil as sh #for file copy
import datetime as dt
import sys 

#Kevin Faterkowski

#https://stackoverflow.com/questions/54244230/async-keyword-in-module-name-is-preventing-import
#https://github.com/riptideio/pymodbus/issues/320
#NOTE: have to install with "pip install python3-ldap"
#NOTE: had to rename async.py to _async.py in C:\Users\00616891\AppData\Local\Programs\Python\Python39-32\Lib\site-packages\ldap3\strategy AND
#  the reference to it was in C:\Users\00616891\AppData\Local\Programs\Python\Python39-32\Lib\site-packages\ldap3\core
#NOTE: had to also install 'openpyxl'
#NOTE: the solution relative to ldap3 was "pip uninstall ldap3" followed by "pip install ldap3"
#need to follow steps here for LDAP https://answers.microsoft.com/en-us/windows/forum/all/microsoft-visual-c-140/6f0726e2-6c32-4719-9fe5-aa68b5ad8e6d

#12/10/2022 - added a authentication loop
#12/25/2022 - moved any network/domain strings to constants

#############################################################################################
#Constants to populate before execution:
#############################################################################################
## LDAP_search_base: populate with LDAP search base - ou: organizational unit, o: organization, c: country, dc: domain
##Example: "OU=<organizational unit/container>,DC=<domain>,DC=com"
##AD_domain: populate as "<domain>.com"
##logon_domainL populate as "<domain>\\"
#############################################################################################
LDAP_search_base = ""
AD_domain = ""
logon_domain = ""
#############################################################################################

def read_excel( path ):
    try:
        df = pd.read_excel(path,header=None) #set this up to run with/without headers
        #reverse columns order by https://www.marsja.se/six-ways-to-reverse-pandas-dataframe/
        columns = df.columns.tolist()
        columns = columns[::-1]
        df = df[columns]
        
        for columns in df:  #"columns" is just an interator/integer and not the object itself..
        #Each column in a DataFrame is a Pandas Series. As a single column is selected, the returned object is a Series.
        #you can use double brackets in order to return a DataFrame datatype, rather than a pandas Series

            if df[columns].str.find('@').any():  
                df[columns].name = 'email'    #if any of the characters are an @, then we have found the email column
                df = pd.DataFrame(df[columns])  #remove all other columns from the dataframe
                #https://thispointer.com/drop-first-row-of-pandas-dataframe-3-ways/                
                if( "@" not in str(df.iloc[0:1]) ): #the first element in this DataFrame is NOT an email
                    df = df.iloc[1: ] #"iloc" takes a slice of a dataframe, the first record is not a header record so delete it
                    break  
        return df

    except FileNotFoundError as e:
        print("this the FNF error", e)
        sys.exit() #terminate the whole program
    except IOError as io:
        print("this the IO error: ", io)
        sys.exit() #terminate the whole program
        
def logon_AD():
    
    attempts = 3    #counter for failed login attempts
    
    #while True:
    while attempts > 0:
        
        #formulate a dynamic input message based on attempts
        if attempts == 3:    #first time prompting
            username_msg = "please enter your logon - "
            pw_msg = "please enter your network password - "
        else:
            username_msg = "please enter your logon - " + str(attempts) + " attempts remaining: "
            pw_msg = "please enter your network password - " + str(attempts) + " attempts remaining: "
            
        #prompt for username and pw and remove excess spaces, and relay attempt count to user:
        username = input(username_msg).strip()
        password = input(pw_msg).strip()
        
        #test for nulls
        if ((username) and (password)):
            conn = connect_AD(username, password)
        else:
            print("Both the logon, password cannot be blank")
            continue
        
        if (conn):
            break     #LDAP connection is good, exit the loop
        else:
            attempts -= 1

    return conn
    
#end logon_AD()

def connect_AD( username, password ):
    
    try:

        #tls_configuration = Tls(validate=ssl.CERT_NONE, version=ssl.PROTOCOL_TLSv1)

        
        server = Server(AD_domain, get_info=ALL)
                               
        #note: if you do not supply your credentials and the Connection instantiation correctly, you will only achieve an anonymous bind:
        #https://stackoverflow.com/questions/25981653/retrieving-attributes-from-ad-python-3
        #https://www.watchguard.com/help/docs/help-center/en-US/Content/en-US/Fireware/authentication/find_ad_search_base_c.html
        #https://social.technet.microsoft.com/Forums/Lync/en-US/3c137047-23c6-41e0-8a34-0b32423c7403/how-to-find-the-ldap-servers-in-a-domain?forum=winserverDS
        #https://ldap3.readthedocs.io/en/latest/bind.html

        conn = Connection( server,
                          user = logon_domain + username,
                          password=password,
                          auto_bind=True )
        

        # perform the Bind operation
        if not conn.bind():
            print('error in binding connection to Active Directory', conn.result)    

        return conn
     
    #except LDAPException as e:
    #    print("Active Directory connection issue: ", e)
    except LDAPCommunicationError as e:
        print("Active Directory connection issue: ", e)
        return
    except LDAPBindError as e:
        print("Bind issue: ", e)
        return
    return
    #end function connect_AD()

def pull_logons( connection, logons ):
        #received help from https://www.youtube.com/watch?v=fhQE342ZTrk&t=463s

        #NOTE: you must ensure that you are pointing to the correct LDAP search base - ou: organizational unit, o: organization, c: country, dc: domain

    logons['cn'] = None #append a new null column onto the dataframe
        
    bad_chars = ['(', ")"]      #initialize a string for stripping parenthesis from LDAP result

    try:
        # iterate through each dataframe row and select 'email' and use it as an LDAP search parameter
        with connection as conn:
            for logon in logons.index:

                conn.search(LDAP_search_base,"(mail=%s)"%(logons['email'][logon]),attributes = ['SamAccountName']) #use the % variable insertion to string trick
                
                if (conn.entries): #successful result from search

                    temp_cn = ''.join(i for i in conn.entries[0]['SamAccountName'] if not i in bad_chars) #This is a trick to remove "bad" chars from a string using .join
                    logons['cn'][logon] = temp_cn
                    #logons['cn'][logon] = print(conn.entries[0]['SamAccountName'])

        return logons
    
    except LDAPCommunicationError as e:
        print("connection issue: ", e)
        return
    except LDAPBindError as e:
        print("Bind issue: ", e)
        return
    except ldap3.LDAPError as e:
        print("connection issue! ", e)
    
    #end function

def Tokenize_Paths( report_paths ):

    report_tokens = [] #storing an empty list
    
    #Each key-value pair in a Dictionary is separated by a comma, whereas each key is separated by a colon
    #report_tokens.append({"server":"aw1dsapwvwp01","berver":"zzzzdsapwvwp01"})
    #report_tokens.append({"server":"aw1dsapwvwp01","berver":"zzzzdsapwvwp01"})
    #report_tokens.append({"server":"aw1dsapwvwp01","berver":"vwp01"})
    #print("report_tokens[2][""berver""] ", report_tokens[2]["berver"])

    for report_path in report_paths:
        report_path = report_path.replace("//","") #the .replace() method has to return a string for you to work with
        report_tokens.append( report_path.split("/") )
    
    return report_tokens
    #end function


# def Get_Reports():

#     #https://stackoverflow.com/questions/1406145/how-do-i-get-rid-of-python-tkinter-root-window
#     #root = tkinter.Tk()
#     root = tk.Tk()
#     root.withdraw()
#     root.wm_attributes('-topmost', 1)
    
#     #askopenfilenames() returns a tuple of strings, not a string
#     report_path = askopenfilenames(
#         initialdir='\\\\AW1DSAPWVWD01\\c$\\Peloton\\My Reports\\WV10 My Reports\\',
#         filetypes=[
#             ("AFR", "*.afr"), ("AFQ", "*.afq"),
#             ("AFR", "*.afr"), ("AFQ", "*.afq"), ("AFSQL", "*.afsql"), ("AFM", "*.afm"), ("AFMXL", "*.afmxl"),
#             ("AFRXL", "*.afrxl"), ("XLT", "*.xlt"), ("XLTM", "*.xltm"), ("XLTX", "*.xltx")
#             ]
#         )
    
#     if (report_path):
#         return report_path 
#     else:
#         print("no report file was selected")
#     sys.exit()
    
#     return
#     #end function

def Get_Filenames( Report_type ):

    #https://stackoverflow.com/questions/1406145/how-do-i-get-rid-of-python-tkinter-root-window
    #root = tk.Tk()
    root = tk.Tk()
    root.wm_attributes('-topmost', 1)
    root.withdraw()

    if Report_type == "Report":
        #askopenfilenames() returns a tuple of strings, not a string
        file_path = askopenfilenames(
            title='SELECT THE REPORT FILES TO BE DISTRIBUTED',
            #initialdir='\\\\AW1DSAPWVWP01\\c$\\Peloton\\My Reports\\WV10 My Reports\\',
            initialdir='\\\\AW1DSAPWVWP01\\WV10 My Reports\\',
            filetypes=[
                ("Excel Pivot files", ".xltx .afmxl .xltm .xlt .afrxl .afmxl"),
                ("Multi well PDF", ".afm"),
                ("Single well PDF", ".afr"),
                ("Query files"," "".afq" ".afsql"),
                ]
            )
    elif Report_type == "Distribution List":
        file_path = askopenfilename(
            title='SELECT THE EMAIL LIST SPREADSHEET',
            initialdir='\\C:\\',
            filetypes=[
                ("XLSX", "*.xlsx"), ("XLS", "*.xls"), ("CSV", "*.csv")
                ]
            )
    
    if (file_path):
        return file_path 
    else:
        print("no report file was selected")
    sys.exit()
    
    return
    #end function


def Distribute_Report( logons, #dataFrame
                       report_tokens #2D list of strings
                       ):

    logon_arr = logons.to_numpy()  #convert everything to numpy
    source_tokens = np.array(report_tokens) #convert everything to numpy

    with open("MyReportsCopyLog.txt", mode="a") as file:
    #with open("C:\\REPORTS COPY TOOLS\\MyReportsCopyLog.txt", mode="a") as file:
        move_date = dt.datetime.now()
        file.write(move_date.strftime("%c") + "\n") #no need to use Python string formatting for this file write action
        
        #loc gets rows (and/or columns) with particular labels.
        #iloc gets rows (and/or columns) WITH INTEGER ONLY locations.

        #for each logon
        for logon in logon_arr: #iterate through 2d numpy array
     
            if(logon[1]): #a logon for the individual email WAS found in Active Directory...

                for path_tokens in source_tokens:  #iterate through 2d numpy array
                    
                    source_path = "\\\\"+"\\".join(path_tokens)  #for each file token path, save the source path
                    dest_tokens = np.array(path_tokens) #do this to preserve the source path for future loop iterations
                    dest_tokens[2] = logon[1]  #insert login into what will become the destination path
                    dest_tokens = dest_tokens[0:dest_tokens.shape[0]-1]  #delete the last column from the numpy array
                    dest_path = "\\\\"+"\\".join(dest_tokens) #create destination path

                    try: 
                        
                        sh.copy(source_path, dest_path) #use .copy() instead of copy2 so that metatdata is NOT copied
                        print(source_path, " COPIED TO: ", "\n", dest_path, "\n")
                        file.write("%s COPIED TO: %s \n\n" %(source_path, dest_path)) #this is called Python string formatting..

                    # If source and destination are same
                    except sh.SameFileError:
                        print("Source and destination represents the same file.")
                     
                    # If destination is a directory.
                    except IsADirectoryError:
                        print("Destination is a directory.")
                     
                    # If there is any permission issue
                    except PermissionError:
                        print("Permission denied.")
                     
                    #For other errors
                    except:
                        print("Error occurred while copying file.")

            else:
                print(logon[0], " was not found.  Report file copy did not happen.  \n\n")
                file.write("%s was not found.  Report file copy did not happen.  \n\n" %(logon[0])) #this is called Python string formatting..
                
    return 
    #end function

def main():

    report_paths = Get_Filenames( "Report" )  #file browser function call for report files
    excel_path = Get_Filenames( "Distribution List" )   #file browser function call for DLs
    logons = read_excel(excel_path) #a DataFrame is returned
    
    conn = logon_AD()
    
    if (conn):
        print("connection to AD successful, continuing...")
        logons = pull_logons( conn, logons ) #a DataFrame is returned
        report_tokens = Tokenize_Paths( report_paths ) #a 2D list of strings is returned
        Distribute_Report( logons, report_tokens ) 
    else:
        print("connection to AD unsuccessful, cannot continue - exiting...")
    
  
    
    #end main() function
    

if __name__ == "__main__":
    main()
    input('Press Enter to Exit...') #this is a trick to keep the window open in Win
