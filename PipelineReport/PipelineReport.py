import os
import tkinter
from tkinter import *
from tkinter import filedialog
import subprocess
import sys
import csv
import xlrd
import math
import datetime
import win32com.client
window = tkinter.Tk() # create window
def pipelineCallBack():
    #prompt for directory with xls reports
    directory = filedialog.askdirectory(initialdir='.')
    #Establish time, create timestamp
    whattime = datetime.datetime.now() 
    timestamp = str(whattime.month)+'-'+str(whattime.day)+'-'+str(whattime.year)+'-'+str(whattime.hour)+'.'+str(whattime.minute)
    #where am i?
    home = os.getcwd() 
    log = home+'\\log\\' # log file directory for trouble shooting
    #create log file
    PL = open(log+'Pipeline-Report-Log-'+str(timestamp)+'.txt','a')  # log
    PL.write('#####################################################\n') # log
    PL.write('Logging Started:'+timestamp+'\n\n') # log
    for file in os.listdir(directory):
        #where am i?
        home = os.getcwd() 
        helpers = home+'\\bin\\' # just incase you ever need to call some external apps
        report = home+'\\report\\' # where the eventual report/csv file will be stored
        conf= home+'\\conf\\' # templates
        log = home+'\\log\\' # logging to see what goes wrong, outside of stdout feedback
        #only xls files in this directory
        if file.endswith(".xls"):
          xl_workbook = xlrd.open_workbook(directory+'/'+file) #open the file (if it's an xls file)
          PL.write('Excel Filename:\n'+str(directory+'/'+file)+'\n\n') # log
          sheet_names = xl_workbook.sheet_names() # query xls file for sheet names
          PL.write('Workbook Names:'+str(sheet_names)+'\n\n') # log
          #regex magic
          srv_noxls = file.replace('.xls','')
          srv_dashpattern1 = re.compile('([a-z])\-')
          srv_nodashpattern1= srv_dashpattern1.sub(r"\1,",srv_noxls)
          srv_dashpattern2 = re.compile('([0-9]{2}\-[0-9]{2}\-[0-9]{4})\-')
          srv_nodashpattern2 = srv_dashpattern2.sub(r"\1,",srv_nodashpattern1)
          #if the excel worksheet 'Bill of Materials' is fount - it assumes it is a certain product line
          if sheet_names == ['Bill of Materials']:
              #product line A's worksheet requires me to do some simple equations and removal of alphabetical characters in some columns
              #temp file created
              rpt_tmp = open(report+'report_tmp.csv', 'a')
              sheet0 = xl_workbook.sheet_by_index(0) #First worksheet
              column_values0 = sheet0.col_values(colx=4) # column e values
              # replace the words 'Line Cost' and 'Included' with 0  as a place filler.  Drop the Dollar sign
              strpped_colvals0 = str(str(str(str(column_values0).replace('Included','0')).replace('Line Cost','0')).replace('$','')).replace('\'\',','\'0\',')
              # remove commas on each line (ie '1,000' - '2,532') as it will screw with a CSV file.
              dropped_comma_colvals0 = re.compile('([0-9]),')
              stripped_colvals0 = dropped_comma_colvals0.sub(r"1",strpped_colvals0)
              drop_quotes = str(stripped_colvals0).replace('\'','')
              drop_bracket= str(drop_quotes).replace('[','')
              drop_bracket0 = str(str(drop_bracket).replace(']','')).replace(' ','')
              drop_comma = str(drop_bracket0).replace(',','\n')
              # write all of these changes to a new temp file.
              ProductLineA_tmp = report+'ProductLineA_tmp.txt'
              ProductLineA_shit = open(ProductLineA_tmp,'w')
              ProductLineA_shit.write(str(drop_comma))
              ProductLineA_shit.close()
              # open this file and calculate the total value of the numbers stored on each line
              with open(ProductLineA_tmp) as f: 
                  tot_sum = 0 
                  for i,x in enumerate(f, 1): # add up the values on each line together
                      val = float(x)
                      tot_sum += val # total value of the bill of materials is determined
                  f.close()
                  os.remove(ProductLineA_tmp) #remove temp file
              PL.write('Final CSV output:\n'+str(srv_nodashpattern2)+','+str(tot_sum)+'\n\n') # log
              # write sum to report
              rpt_tmp.write(srv_nodashpattern2+","+str(tot_sum)+"\n") # write the total sum of the temp file to the CSV file
              rpt_tmp.close()
          # ProductLineB determined by unique worksheet names
          if sheet_names == ['Quote', 'Sheet2', 'Sheet3']:
              sheet1 = xl_workbook.sheet_by_index(0) # grab the first worksheet
              column_values1 = sheet1.col_values(colx=10) # column K
              rpt_tmp0 = open(report+'report_tmp.csv', 'a') # open temp report (in append mode)
              rpt_tmp0.write(srv_nodashpattern2+","+str(column_values1[-1])+"\n") # append the last value of column k at the end of each line of the csv file
              rpt_tmp0.close() 
    reportname = open(report+'UserName-PipelineReport-'+timestamp+'.csv','w')
    tempread = open(report+'report_tmp.csv')
    with tempread as f: # loop through csv file and attempt to guess the person's email address at your company
        reader = csv.reader(f, delimiter=',') # read csv file
        reportname.write("Account Manager,Company Name,Date,Model,Price,Email\n") # create column headers for csv file
        for row in reader: # loop dee loop!
            emailpattern = re.sub("([a-z])([A-Z])","\g<1>.\g<2>", row[0]) # regex - detect camel case and insert a period inbetween small letters followed by a capital
            noemailpattern= emailpattern+"@companyname.com"  # assuming the firstname.lastname@companyname.com syntax
            reportname.write(','.join(row)+','+noemailpattern+"\n") # write this assumption to the end of each line 
        reportname.close()
        tempread.close()
    os.remove(report+"report_tmp.csv") # delete temp file
    PL.close() # log
def accountmgrCallBack(): # Email Button Action
    # establish paths. where am i?!
    home = os.getcwd() # script root
    helpers = home+'\\bin\\' # helpers
    report = home+'\\report\\' # output
    conf= home+'\\conf\\' # templates
    log = home+'\\log\\' # logging/troubleshooting
    reportname0 = filedialog.askopenfilename(initialdir=report) # prompt for csv file
    # read the csv file
    csvfile = reportname0
    e = open(reportname0)
    csv_e = csv.reader(e)
    for row in csv_e: # open the CSV file and use the information to send an outlook email based off the information in that line.
        # assign each column in the csv to a variable
        name = row[0] #csv file column name
        firstname0 = re.sub("([a-z])([A-Z])","\g<1> \g<2>", row[0])# regex - put a space between frist and last
        firstname1 = re.compile('([A-Z]{1}[a-z]{1,15} )([A-Z]{1}[a-z]{1,15})') # regex - compile it
        firstname2 = firstname1.sub(r"\g<1>",firstname0) #regex - drop the last name. JohnSmith is now John
        company = row[1] # csv file column name
        bomdate = row[2] # csv file column name
        model = row[3] # csv file column name
        price = row[4] # csv file column name
        email = row[5] # csv file column name
        # spawn an outlook process/object
        olMailItem = 0x0
        obj = win32com.client.Dispatch("Outlook.Application")
        # create and send email
        newMail = obj.CreateItem(olMailItem)
        newMail.To = str(email)
        newMail.Cc = "someemail@companyname.com" # just incase you want to send it to another inbox
        newMail.Subject = "Sales Follow Up From "+str(company)+" on "+str(bomdate) # email subject line
        newMail.Body = "Your Message Here" # message content... apply vars/strings as necessary.
        newMail.Send() # SEND!
    e.close()
# fill window contents
title = Label(window, text="Pipeline Report Helper")  
pipeline = tkinter.Button(window, text="Create Pipeline Report", width=100,  command = pipelineCallBack)
accountmgr = tkinter.Button(window, text="Email Account Managers", width=100, command = accountmgrCallBack)
# arrange window 
title.pack()
pipeline.pack()
accountmgr.pack()
window.mainloop()
