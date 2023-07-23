'''

To automate daily reports, I will:

    1. Update data from the SFTP server to the local SQL database.
    2. Execute SQL queries to generate Excel reports on a daily basis.
    3. Refresh Power Query reports that are linked to the SQL reports I have implemented.
    4. Email each report to the appropriate individual.

Developing this project saved my department 20 hours of work per week.

'''



#_________________________________________________________________________#
#_________________________________________________________________________#
#______________________________Packages___________________________________#

# Import Packages:

import schedule
import time
import numpy as np
import pandas as pd    
import os
import glob
import pyodbc
import shutil
import win32com.client as win32
import paramiko
import sys
import datetime
import time
from pathlib import Path
from PIL import ImageGrab
from datetime import timedelta
from datetime import date
from datetime import datetime

# Print some words:

print("Packages are imported.")


#_________________________________________________________________________#
#_________________________________________________________________________#
#________________________________Dates____________________________________#

'''

Note:
    Excel dates are represented as the number of days since January 1, 1900.
To convert a date in the usual format to the Excel format,
I will define a function that calculates the difference between the two dates.


As part of my project, I need to specify the dates of:
    (today, yesterday, the day before yesterday and the begin of month).

This will allow me to keep track of all four dates.

'''

# First, specify dates in a regular format:

xl_date = date(1899, 12, 30)
today_date = date.today()
yesterday_date = date.today() - timedelta(days = 1)
first_yesterday = date.today() - timedelta(days = 2)
month_begin = yesterday_date.replace(day=1)

# Second, convert the dates to string:

str_today = datetime.strftime(today_date , '%m-%d-%Y')
str_yesterday = datetime.strftime(yesterday_date , '%m-%d-%Y')
str_first_yesterday = datetime.strftime(first_yesterday , '%m-%d-%Y')
str_month_begin = datetime.strftime(month_begin , '%m-%d-%Y')

# Print some words to verify the results:

print(f'Today is {str_today}')
print(f'Yesterday is {str_yesterday}')
print(f'First yesterday is {str_first_yesterday}')
print(f'Begin of the month is {str_month_begin}') 

# Then, get the dates in Excel format:

today_xl = today_date - xl_date
today_xl = today_xl.days
yesterday_xl = today_xl -1
first_yesterday_xl = yesterday_xl -1
month_begin_xl = month_begin - xl_date
month_begin_xl = month_begin_xl.days

# Finally, print some words to verify the results:


print(f'Today in Excel is {today_xl}')
print(f'Yesterday in Excel is {yesterday_xl}')
print(f'First yesterday in Excel is {first_yesterday_xl}')
print(f'Begin of the month in Excel is {month_begin_xl}')

#_________________________________________________________________________#
#_____________________________Set_Jobs____________________________________#
#______________________________ Job_1 ____________________________________#
#_________________________Daily_Throughput________________________________#

# Define Daily Throughput Report:

def Daily_Throughput(): # This function provides a set of instructions for creating periodic reports. #

#_________________________________________________________________________#
#__________________________Job_1-Section_1 _______________________________#
#_________________________Transfer Files__________________________________#


    '''

In my project, I need to import daily sales files from the SFTP server to my local machine.

The filenames have:
    constant part
and
    variable part
        the date of the previous day.

I will create a dictionary with:
    Key
        the file path and constant part of the filename.
    Value
        the new file path on the local machine.

I will then iterate through the dictionary, adding the date variable as text and the ".xlsx" extension to each key and value.

This will allow me to move the files from the remote machine to the local machine in a for loop.

    '''

# First, Create an SSHClient:

    client = paramiko.SSHClient() # This class creates a client that requires remote server credentials. #

# Define remote server credentials:

    host = "host_name"
    username = "your_username"
    password = "your_password"
    port = 22

    key_filenames = False

# You can also create the client directly with this line of code:

    client.set_missing_host_key_policy(paramiko.AutoAddPolicy())

# Connect the client:

    client.connect(hostname=host,port=port,username=username,password=password)

# Create an SFTP client object:

    ftp = client.open_sftp()

#_________________________________________________________________________#

# Second, download files from the remote server:


# 1-Create the dictionary:

    serv_dict = {r"remote_path1":r"local_path1",
                r"remote_path2":r"local_path2",
                r"remote_path3":r"local_path3",
                r"remote_path4":r"local_path4"}

# 2-Use the variable for yesterday's date, which I defined in a previous section, in the for loop:

    for x,y in serv_dict.items():
        try:
            ftp.get(x + str_yesterday + ".xlsx", y + str_yesterday + ".xlsx") # Moving files to the local machien. #

            print(x[8:] + str_yesterday + " .xlsx has been moved.") # Verify which files have been moved. #
            x += x
        except:
            if FileNotFoundError: "[Errno 2] No such file" # Some files may not exist, so I will ignore the error. #

            print(x[8:] + str_yesterday + " .xlsx is not exist.") # Verify if that occurs. #

            os.remove(y + str_yesterday + ".xlsx") # Should that occur, the file will be created anyway, so I will delete it. #
            
            break

#_________________________________________________________________________#
#__________________________Job_1-Section_2 _______________________________#
#_____________________________Dataframe___________________________________#

# Define inserting excel files from folder into dataframe:

    def import_files(path): # This function provides a set of instructions for inserting excel files from folder into dataframe. #

    # 1-Print start time:

        start_time = datetime.now()

        print("Begin:",start_time)


    # 2-Import files:

        path = path
        files = glob.glob(os.path.join(path, "*.xlsx")) # Import all files with '.xlsx' from a single folder. #

        filenames = [file for file in files] # Put file names in a list. #

        '''

The imported files have the same fields, but they have been split into several files, each containing a million records.

To make it easier to work with the data, I will combine these files into one dataframe.

        '''

        df = pd.concat([pd.read_excel(file) for file in filenames], ignore_index=True)


    # 3-Print file names and number:

        num = len(filenames)

        for file in filenames:
            print("Imported:", file) # Check file names. #

        print("Total Number of imported files:", num)



# 4-Remove files to avoid repeating the process:
 
        for f in files:
            os.remove(f)


    # 5-Print end time:

        end_time = datetime.now()
        print("End:", end_time)


    # 6-Print Total time:

        total_time = end_time - start_time
        print("Total Time:", total_time)
    
        return df

#_________________________________________________________________________#


# Insert excel files from folder into dataframe #

    '''

I will import the sales files that I have transferred from the remote machine to the local machine
into a dataframe
using the import_files function.

    '''

    path = r"folder_path"

    df = import_files(path)

#_________________________________________________________________________#
#__________________________Job_1-Section_3 _______________________________#
#______________________________SQL TBL____________________________________#


    '''
In this section:
    First,
        I will use data manipulation techniques to fit the dataframe to the table in my local SQL database.
    Then,
        I will connect to SQL Server.
    Finally,
        I will insert the data into my SQL table.

    '''
#_________________________________________________________________________#

# Manipulation

# 1-Column names:

    df = df.rename(
        columns = {
            
            "حقل_1":"Field_1",
            "حقل_2":"Field_2",
            "حقل_3":"Field_3",
            "حقل_4":"Field_4",
            "حقل_5":"Field_5",
            "حقل_6":"Field_6",
            "حقل_7":"Field_7",
            "حقل_8":"Field_8",
            "حقل_9":"Field_9",
            "حقل_10":"Field_10"
    
        }
    )


# 2-Column types:

    df[("Field_1")] = df[("Field_1")].astype(str)

    df = df.astype(
        columns = {
            "Field_1":str,
            "Field_2":int
        }
    )


# 3-Null values:

    df.replace({np.inf: np.nan, -np.inf: np.nan}, inplace=True)
    df = df.fillna(0)

#_________________________________________________________________________#


# Connect to SQL Server:

    conn = (
        r'DRIVER={SQL Server};'
        r'SERVER=Your-Server-Name;'
        r'DATABASE=Your-Database-Name;'
        r'Trusted_Connection=yes;'
    )
    cnxn = pyodbc.connect(conn)

    cursor = cnxn.cursor()

#_________________________________________________________________________#


# Insert dataframe into SQL table:

    for row in df.itertuples():
        cursor.execute('''
                    INSERT INTO Your_TBL (Field_1,
                    Field_2,
                    Field_3,
                    Field_4,
                    Field_5,
                    Field_6,
                    Field_7,
                    Field_8,
                    Field_9,
                    Field_10)
                    VALUES (?,?,?,?,?,?,?,?,?,?)
                    ''',
                row.Field_1, 
                row.Field_2,
                row.Field_3,
                row.Field_4,
                row.Field_5,
                row.Field_6,
                row.Field_7,
                row.Field_8,
                row.Field_9,
                row.Field_10,
                    )
    cnxn.commit()

#_________________________________________________________________________#
#__________________________Job_1-Section_4 _______________________________#
#___________________________SQL Queries___________________________________#


# SQL-Query Daily_Throughput:

    daily_throughput = pd.read_sql(

        '''
        SELECT
            Field_1,
            Field_2,
            Field_3,
            Field_4,
            SUM(Field_5) AS Sales_Out,
            COUNT(Field_5) AS Transactions_Count
        FROM
            Your_TBL
        WHERE
            Date_Field >''' + str(yesterday_xl) + '''
        AND
            Date_Field <=''' + str(today_xl) + '''
        GROUP BY
            Field_1,
            Field_2,
            Field_3,
            Field_4;
        ''', cnxn
    )

# Export an excel file:

    daily_throughput.to_excel(r"path" + str(yesterday_xl) + ".xlsx", index=False)

#_________________________________________________________________________#

# SQL-Query MTD_Throughput:

    mtd_throughput = pd.read_sql(
        
        '''
        WITH
        MTD_Throughput AS(
            SELECT
                Field_1 AS Field_1,
                Field_2 AS Field_2,
                Field_3 AS Field_3,
                Field_4 AS Field_4,
                SUM(Field_5) AS MTD_Sales,
                COUNT(Field_5) AS MTD_Transactions
            FROM
                Your_TBL
            WHERE
                Date_Field >=''' + str(month_begin_xl) + '''
            GROUP BY
                Field_1,
                Field_2,
                Field_3,
                Field_4),
        Daily_Throughput AS(
            SELECT
                Field_1 AS Field_1,
                Field_2 AS Field_2,
                Field_3 AS Field_3,
                Field_4 AS Field_4,
                SUM(Field_5) AS Daily_Sales,
                COUNT(Field_5) AS Daily_Transactions
            FROM
                Your_TBL
            WHERE
                Date_Field >=''' + str(yesterday_xl) + ''' 
            AND
                Date_Field <=''' + str(today_xl) + ''' 
            GROUP BY
                Field_1,
                Field_2,
                Field_3,
                Field_4)
        SELECT
            l.Field_1,
            l.Field_2,
            l.Field_3,
            l.Field_4,
            MTD_Sales,
            MTD_Transactions,
            Daily_Sales,
            Daily_Transactions
        FROM
            MTD_Throughput l
        LEFT JOIN
            Daily_Throughput r
        ON
            l.Field_1 = r.Field_1
        AND
            l.Field_2 = r.Field_2
        AND
            l.Field_3 = r.Field_3
        AND
            l.Field_4 = r.Field_4;
        ''', cnxn)

# Export excel files:

    '''

As the month progresses, this query will likely exceed
    one million records
and may reach
    three million.
These records exceed the capacity of a single Excel file, so I will handle them as follows:

    '''
    rows_number = len(mtd_throughput) # Find out the number of records. #
    print(rows_number)

    if rows_number <= 1000000: # Then, export the DataFrame to a single file. #

        mtd_throughput[0:1000000].to_excel(
            r'path\filename' + str_yesterday + '_Part1.xlsx',
            index=False
        )

    elif rows_number > 1000000 and rows_number <= 2000000: # Then, split the DataFrame to two files. #

        mtd_throughput[0:1000000].to_excel(
            r'path\filename' + str_yesterday + '_Part1.xlsx',
            index=False
        )
        mtd_throughput[1000000:2000000].to_excel(
            r'path\filename' + str_yesterday + '_Part2.xlsx',
            index=False
        )

    else: # Given that the records will not exceed three million, we can split the DataFrame to three files. #

        mtd_throughput[0:1000000].to_excel(
            r'path\filename' + str_yesterday + '_Part1.xlsx',
            index=False
        )
        mtd_throughput[1000000:2000000].to_excel(
            r'path\filename' + str_yesterday + '_Part2.xlsx',
            index=False
        )
        mtd_throughput[2000000:3000000].to_excel(
            r'path\filename' + str_yesterday + '_Part3.xlsx',
            index=False
        )
    
        '''

    The size of the files exceeds 20 MB
        so:
            I will upload them to a shared folder where team members can access them.
            Since
                this report is for sales for the month to date:
            I will discard the old files.
        
        '''
    # Remove old files:

    old_files = [
        r'path\filename' + str_first_yesterday + '_Part1.xlsx',
        r'path\filename' + str_first_yesterday + '_Part2.xlsx',
        r'path\filename' + str_first_yesterday + '_Part3.xlsx'
    ]

    try:
        for file in old_files:
            os.remove(file)
    except:
        FileNotFoundError
    pass

#_________________________________________________________________________#
#__________________________Job_1-Section_5 _______________________________#
#___________________________Power Query___________________________________#

# Refresh power query files:
    
    '''
    
Since there are
    reports based on Power Query
and
    it was linked with the SQL reports I have implemented,
I will refresh them. 

    '''
 
    files = [

        r"path_1",
        r"path_2",
        r"path_3"

    ] # List all files to refresh. #

    excel = win32.gencache.EnsureDispatch('Excel.Application') # Connect With Excel. #
 
    for f in files: # Iterate over the files and process them in order.

        wb = excel.Workbooks.Open(f) # Open file. #
 
        excel.Visible = True # Show file. #
 
        wb.RefreshAll() # Refresh all data connections. #

        time.sleep(10) # Wait for the update before saving
 
        wb.Close(SaveChanges=True) # Save changes. #
 
        excel.Application.Quit() # Quit file. #

    f += f # Process the files one at a time. #

#_________________________________________________________________________#
#_____________________________Set_Jobs____________________________________#
#______________________________ Job_2 ____________________________________#
#_______________________________Mails_____________________________________#


def Mails():

#_________________________________________________________________________#
#__________________________Job_2-Section_1 _______________________________#
#_________________________Def Send E-Mails________________________________#

# Define sending emails:

    def send_email(sub,to,cc,body):

    
        today = date.today()
        yesterday = today - timedelta(days = 1)
        date_title = yesterday.strftime('%#d %b %Y') # So that I can include it in the subject of some mails. #

        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.Subject = sub + date_title
        mail.To = to
        mail.CC = cc
        mail.HTMLBody = body

        '''

         I will configure the function to attach:
            an Excel file,
            a first image,
            and a second image
        if they exist.

        '''
        try:
            if att != "None":
                os.chdir(dir)
                mail.Attachments.Add(os.getcwd() + att)  
        except:
            pass


        try:
            if image_path != "None":
                mail.Attachments.Add(Source=image_path)
        except:
            pass

        try:
            if image_path2 != "None":
                mail.Attachments.Add(Source=image_path2)
        except:
            pass

        return mail.Send()

#_________________________________________________________________________#
#__________________________Job_2-Section_2 _______________________________#
#______________________________Mails______________________________________#
#______________________________Mail_1_____________________________________#


# Send mail with attached excel file:

    sub = 'Your Subject '
    to = 'xxx@aaa.com; yyy@aaa.com'
    cc = 'zzz@aaa.com'
    body = r'''
    Good morning dears!<br><br>
    Attached is the ... Pleas check it.<br><br>
    BR,<br>
    Your signature<br>
    Your department
    '''
    dir = r'path'
    att = r'\filename_' + str(yesterday_xl) + '.xlsx'
    image_path = "None"
    image_path2 = "None"

    send_email(sub,to,cc,body)

#_________________________________________________________________________#
#__________________________Job_2-Section_2 _______________________________#
#______________________________Mails______________________________________#
#______________________________Mail_2_____________________________________#

# Send mail with attached copy of table as an image:

    excel_path = r"excel_path"
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Open(excel_path)
    ws = wb.Worksheets(2)

    win32c = win32.constants
    ws.Range("A1:F30").CopyPicture(Format=win32c.xlBitmap) # Replace the table range with your own. # 
    img = ImageGrab.grabclipboard()
    image_path = r"image_path"
    image_path2 = "None"
    img.save(image_path)
    wb.Close()



    
    folder = 'shared_folder_path'

    body = r'''
    Good morning!<br><br>
    Please note that the ... is updaetd. follow this link:
    <br><br>" + folder + "<br><br>
    And here is a summary:
    <br><br> <img src=ImageName.png><br><br>
    BR,<br>
    Your signature<br>
    Your department<br>
    '''

    sub = 'Your Subject '
    to = 'xxx@aaa.com; yyy@aaa.com'
    cc = 'zzz@aaa.com'
    body = (body)
    att = "None"

    send_email(sub,to,cc,body)

#_________________________________________________________________________#
#__________________________Job_2-Section_2 _______________________________#
#______________________________Mails______________________________________#
#______________________________Mail_1_____________________________________#

# Send mail with attached excel file and copy of tow tables as images:

    excel_path = r"excel_path"
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Open(excel_path)
    ws = wb.Worksheets(5)

    win32c = win32.constants
    ws.Range("A1:D7").CopyPicture(Format=win32c.xlBitmap) # Replace the table range with your own. # 
    img = ImageGrab.grabclipboard()
    image_path = r"image_path"
    img.save(image_path)
    wb.Close()

    excel_path = r"excel_path"
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Open(excel_path)
    ws = wb.Worksheets(1)

    win32c = win32.constants
    ws.Range("A1:E22").CopyPicture(Format=win32c.xlBitmap) # Replace the table range with your own. # 
    img = ImageGrab.grabclipboard()
    image_path2 = r"image_path2"
    img.save(image_path2)
    wb.Close()

# Send email:

    body = r'''
    Good morning dears!<br><br>
    Attached is ...:
    <br><br> <img src=ImageName.png><br><br>
    <br><br> <img src=ImageName2.png><br><br>
    BR,<br>
    Your signature<br>
    Your department<br>
    '''
    dir = r"path"
    att = r"\filename"
    sub = 'Your Subject '
    to = 'xxx@aaa.com; yyy@aaa.com'
    cc = 'zzz@aaa.com'
    body = (body)
    send_email(sub,to,cc,body)

#_________________________________________________________________________#
#_________________________________________________________________________#
#_____________________________Run_Jobs____________________________________#


schedule.every().day.at("05:00").do(Daily_Throughput) # Schedule the Daily_Throughput function to execute at 05:00 AM. #  
schedule.every().day.at("08:00").do(Mails) # Schedule the Mails function to execute at 08:00 AM. #

while True:
    schedule.run_pending()
    time.sleep(1)

#_________________________________________________________________________#