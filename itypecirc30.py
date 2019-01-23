# -*- coding: utf-8 -*-
"""
Created on Wed Oct  3 10:01:09 2018

@author: nallen
"""

# -*- coding: utf-8 -*-
"""
Created on Thu Aug  2 14:32:38 2018

@author: nallen
"""

# -*- coding: utf-8 -*-
#!/usr/bin/python2.7
#
# Recreating Web Managment reports for call no circ stats
# 
# Use XlsxWriter to create spreadsheet from SQL Query
# 
#

import psycopg2
import xlsxwriter
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
from datetime import datetime



excelfile =  'itypecirc.xlsx'



#Set variables for email

emailhost = ''
emailport = '25'
emailsubject = 'itypeCircStats'
emailmessage = '''Here are the  Circ stats for Main, Byram and Cos Cob.'''
emailfrom = ''
emailto = []

try:
    conn = psycopg2.connect("dbname= user= host= port= password= sslmodee")
except psycopg2.Error as e:
    print ("Unable to connect to database: " + str(e))
    
cursor = conn.cursor()
cursor.execute(open("itypeCirc30.sql","r",).read())
rows = cursor.fetchall()
conn.close()


workbook = xlsxwriter.Workbook(excelfile, {'remove_timezone': True})

worksheet = workbook.add_worksheet('Total')

worksheet.set_landscape()
worksheet.hide_gridlines(0)



eformat= workbook.add_format({'text_wrap': True, 'valign': 'top' , 'num_format': 'mm/dd/yy'})
eformatlabel= workbook.add_format({'text_wrap': False, 'valign': 'top', 'bold': True})

worksheet.set_header('CircStats')

#TOTALS 

worksheet.set_column(0,0,29)
worksheet.set_column(1,1,10)
worksheet.set_column(2,2,10)
worksheet.set_column(3,3,10)


worksheet.set_header('CircStats')

worksheet.write(0,1,'Main', eformatlabel)
worksheet.write(0,2,'Byram', eformatlabel)
worksheet.write(0,3,'Cos Cob', eformatlabel)
worksheet.write(1,0,'Books', eformatlabel)
worksheet.write(2,0,'Periodicals Digital and Print', eformatlabel)
worksheet.write(3,0,'Music CDs', eformatlabel)
worksheet.write(4,0,'Music Downloadable', eformatlabel)
worksheet.write(5,0,'Audiobook CDs', eformatlabel)
worksheet.write(6,0,'Audiobook Downloadable', eformatlabel)
worksheet.write(7,0,'Video Downloadable', eformatlabel)
worksheet.write(8,0,'DVDs', eformatlabel)
worksheet.write(9,0,'Games', eformatlabel)
worksheet.write(10,0,'Lending Art', eformatlabel)
worksheet.write(11,0,'e-Books', eformatlabel)
worksheet.write(12,0,'Juv Books', eformatlabel)
worksheet.write(13,0,'Juv Music CDs', eformatlabel)
worksheet.write(14,0,'Juv AudioBook', eformatlabel)
worksheet.write(15,0,'Juv DVD', eformatlabel)
worksheet.write(16,0,'Juv Video Game', eformatlabel)
worksheet.write(17,0,'', eformatlabel)
worksheet.write(18,0,'Items Checked-In', eformatlabel)
worksheet.write(19,0,'In House', eformatlabel)
worksheet.write(20,0,'Holds Filled', eformatlabel)
worksheet.write(21,0,'Self Service Checkouts', eformatlabel)
worksheet.write(22,0,'Ct-Card Checkouts', eformatlabel)
worksheet.write(23,0,'Elderly, Nursing & Homebound', eformatlabel)
worksheet.write(24,0,'ILL Loaned', eformatlabel)
worksheet.write(25,0,'ILL Borrowed', eformatlabel)




for rownum, col in enumerate(rows):
 #MAIN TOTALS   
    
    worksheet.write(rownum+18,1,col[58])
    worksheet.write(rownum+19,1,col[59])
    worksheet.write(rownum+20,1,col[60])
    worksheet.write(rownum+21,1,col[61])
    worksheet.write(rownum+22,1,col[62])
    worksheet.write(rownum+23,1,col[63])   
    worksheet.write(rownum+24,1,col[64])
    worksheet.write(rownum+25,1,col[65]) 
    
    worksheet.write(rownum+18,2,col[124])
    worksheet.write(rownum+19,2,col[125])
    worksheet.write(rownum+20,2,col[126])
    worksheet.write(rownum+21,2,col[127])
    worksheet.write(rownum+22,2,col[128])
    worksheet.write(rownum+23,2,col[129])   
    
    worksheet.write(rownum+18,3,col[188])
    worksheet.write(rownum+19,3,col[189])
    worksheet.write(rownum+20,3,col[190])
    worksheet.write(rownum+21,3,col[191])
    worksheet.write(rownum+22,3,col[192])
    worksheet.write(rownum+23,3,col[193])  

    
    worksheet.write_formula('B2', "='Main'!B2+'Main'!C2+'Main'!B7+'Main'!C7+'Main'!B13+'Main'!C13+'Main'!B5+'Main'!C5")
    worksheet.write_formula('B3', "='Main'!B9+'Main'!B8+'Main'!C8")
    worksheet.write_formula('B4', "='Main'!B3+'Main'!C3")
    worksheet.write_formula('B5', "='Main'!B32+'Main'!B33")
    worksheet.write_formula('B6', "='Main'!B15+'Main'!C15")
    worksheet.write_formula('B7', "='Main'!B34+'Main'!B35") 
    worksheet.write_formula('B8', "='Main'!B39+'Main'!B40+'Main'!B41+'Main'!B42")
    worksheet.write_formula('B9', "='Main'!B24+'Main'!C24+'Main'!B22+'Main'!C22")
    worksheet.write_formula('B10', "='Main'!B29+'Main'!C29")
    worksheet.write_formula('B11', "='Main'!B6+'Main'!C6")
    worksheet.write_formula('B12', "='Main'!B36+'Main'!C37+'Main'!C38")
    worksheet.write_formula('B13', "='Main'!B14+'Main'!C14+'Main'!b27+'Main'!C27+'Main'!b28+'Main'!C28")
    worksheet.write_formula('B14', "='Main'!B4+'Main'!C4")
    worksheet.write_formula('B15', "='Main'!B16+'Main'!C16")
    worksheet.write_formula('B16', "='Main'!B23+'Main'!C23")
    worksheet.write_formula('B17', "='Main'!B30+'Main'!C30")
    
#BYRAM TOTALS
    worksheet.write_formula('C2', "='Byram'!B2+'Byram'!C2+'Byram'!B7+'Byram'!C7+'Byram'!B12+'Byram'!C12+'Byram'!B5+'Byram'!C5")
    worksheet.write_formula('C3', "='Byram'!B8+'Byram'!C8")
    worksheet.write_formula('C4', "='Byram'!B3+'Byram'!C3")
    worksheet.write_formula('C5', "='Byram'!B32+'Byram'!B33")
    worksheet.write_formula('C6', "='Byram'!B14+'Byram'!C14")
    worksheet.write_formula('C7', "='Byram'!B34+'Byram'!C35")
    worksheet.write_formula('C8', "='Byram'!B39+'Byram'!B40+'Byram'!B41+'Byram'!B42")
    worksheet.write_formula('C9', "='Byram'!B23+'Byram'!C23+'Byram'!B21+'Byram'!C21")
    worksheet.write_formula('C10', "='Byram'!B28+'Byram'!C28")
    worksheet.write_formula('C11', "='Byram'!B6+'Byram'!C6")
    worksheet.write_formula('C12', "='Byram'!B36+'Byram'!C37+'Byram'!C38")
    worksheet.write_formula('C13', "='Byram'!B13+'Byram'!C13+'Byram'!b27+'Byram'!C27+'Byram'!b26+'Byram'!C26")
    worksheet.write_formula('C14', "='Byram'!B4+'Byram'!C4")
    worksheet.write_formula('C15', "='Byram'!B15+'Byram'!C15")
    worksheet.write_formula('C16', "='Byram'!B22+'Byram'!C22")
    worksheet.write_formula('C17', "='Byram'!B29+'Byram'!C29")


#COS COB TOTALS    
    worksheet.write_formula('D2', "='Cos Cob'!B2+'Cos Cob'!C2+'Cos Cob'!B7+'Cos Cob'!C7+'Cos Cob'!B12+'Cos Cob'!C12+'Cos Cob'!B5+'Cos Cob'!C5+'Cos Cob'!B16+'Cos Cob'!C16")
    worksheet.write_formula('D3', "='Cos Cob'!B8+'Cos Cob'!C8")
    worksheet.write_formula('D4', "='Cos Cob'!B3+'Cos Cob'!C3")
    worksheet.write_formula('D5', "='Cos Cob'!B32+'Cos Cob'!B33")
    worksheet.write_formula('D6', "='Cos Cob'!B14+'Cos Cob'!C14")
    worksheet.write_formula('D7', "='Cos Cob'!B34+'Cos Cob'!C35")
    worksheet.write_formula('D8', "='Cos Cob'!B39+'Cos Cob'!B40+'Cos Cob'!B41+'Cos Cob'!B42")
    worksheet.write_formula('D9', "='Cos Cob'!B23+'Cos Cob'!C23+'Cos Cob'!B21+'Cos Cob'!C21")
    worksheet.write_formula('D10', "='Cos Cob'!B28+'Cos Cob'!C28")
    worksheet.write_formula('D11', "='Cos Cob'!B6+'Cos Cob'!C6")
    worksheet.write_formula('D12', "='Cos Cob'!B36+'Cos Cob'!C37+'Cos Cob'!C38")
    worksheet.write_formula('D13', "='Cos Cob'!B13+'Cos Cob'!C13+'Cos Cob'!b27+'Cos Cob'!C27+'Cos Cob'!b26+'Cos Cob'!C26")
    worksheet.write_formula('D14', "='Cos Cob'!B4+'Cos Cob'!C4")
    worksheet.write_formula('D15', "='Cos Cob'!B15+'Cos Cob'!C15")
    worksheet.write_formula('D16', "='Cos Cob'!B22+'Cos Cob'!C22")
    worksheet.write_formula('D17', "='Cos Cob'!B29+'Cos Cob'!C29")
   

worksheet = workbook.add_worksheet('Main')


worksheet.set_landscape()
worksheet.hide_gridlines(0)



eformat= workbook.add_format({'text_wrap': True, 'valign': 'top' , 'num_format': 'mm/dd/yy'})
eformatlabel= workbook.add_format({'text_wrap': False, 'valign': 'top', 'bold': True})



worksheet.set_column(0,0,29)
worksheet.set_column(1,1,10)
worksheet.set_column(2,2,10)





worksheet.write(0,1,'Checkouts', eformatlabel)
worksheet.write(0,2,'Renewals', eformatlabel)
worksheet.write(1,0,'0 Adult Books', eformatlabel)
worksheet.write(2,0,'2 Adult Music', eformatlabel)
worksheet.write(3,0,'3 Juv Music', eformatlabel)
worksheet.write(4,0,'6 REF Main', eformatlabel)
worksheet.write(5,0,'7 Lending Art', eformatlabel)
worksheet.write(6,0,'8 ILL', eformatlabel)
worksheet.write(7,0,'9 Periodicals', eformatlabel)
worksheet.write(8,0,'9 Digital Periodicals', eformatlabel)
worksheet.write(9,0,'10 Perrot Tote Bag Books', eformatlabel)
worksheet.write(10,0,'16 JUV Spknwrd Cass.', eformatlabel)
worksheet.write(11,0,'17 JUV music Cass.', eformatlabel)
worksheet.write(12,0,'18 Park/Museum Pass', eformatlabel)
worksheet.write(13,0,'20 Juvenile Books', eformatlabel)
worksheet.write(14,0,'22 Adult Spknwrd CDs', eformatlabel)
worksheet.write(15,0,'23 Juv Spnwrd CDs', eformatlabel)
worksheet.write(16,0,'31 Perr-A Express Books', eformatlabel)
worksheet.write(17,0,'37 Perr-J Periodicals', eformatlabel)
worksheet.write(18,0,'38 Perr-A Playaways', eformatlabel)
worksheet.write(19,0,'39 Perr-J Playaways', eformatlabel)
worksheet.write(20,0,'40 E-readers', eformatlabel)
worksheet.write(21,0,'41 Adult Non-Fict DVDs', eformatlabel)
worksheet.write(22,0,'42 DVD Children', eformatlabel)
worksheet.write(23,0,'44 DVD Adult', eformatlabel)
worksheet.write(24,0,'46 Perr-A Express DVDs', eformatlabel)
worksheet.write(25,0,'50 E-Books', eformatlabel)
worksheet.write(26,0,'60 J Books with CD ', eformatlabel)
worksheet.write(27,0,'61 Juv Multimedia Kit', eformatlabel)
worksheet.write(28,0,'62 Adult Game', eformatlabel)
worksheet.write(29,0,'63 Juv Game', eformatlabel)
worksheet.write(30,0,'80 Audbk Download', eformatlabel)
worksheet.write(31,0,'Hoopla Music', eformatlabel)
worksheet.write(32,0,'Naxos Music', eformatlabel)
worksheet.write(33,0,'Audbk Download Overdrive', eformatlabel)
worksheet.write(34,0,'Audbk Download One Click Digital', eformatlabel)
worksheet.write(35,0,'eBook Overdrive', eformatlabel)
worksheet.write(36,0,'eBook Recorded Books', eformatlabel)
worksheet.write(37,0,'eBook Hoopla comics', eformatlabel)
worksheet.write(38,0,'Kanopy Video', eformatlabel)
worksheet.write(39,0,'Hoopla Movies', eformatlabel)
worksheet.write(40,0,'Hoopla TV', eformatlabel)
worksheet.write(41,0,'RB Digital(Qello/Acorn/GC)', eformatlabel)



for rownum, col in enumerate(rows):
 #MAIN Checkout numbers
    worksheet.write(rownum+1,1,col[0])
    worksheet.write(rownum+2,1,col[1])
    worksheet.write(rownum+3,1,col[2])
    worksheet.write(rownum+4,1,col[3])
    worksheet.write(rownum+5,1,col[4])
    worksheet.write(rownum+6,1,col[5])
    worksheet.write(rownum+7,1,col[6])
    worksheet.write(rownum+9,1,col[7])
    worksheet.write(rownum+10,1,col[8]) 
    worksheet.write(rownum+11,1,col[9]) 
    worksheet.write(rownum+12,1,col[10])
    worksheet.write(rownum+13,1,col[11]) 
    worksheet.write(rownum+14,1,col[12])     
    worksheet.write(rownum+15,1,col[13])    
    worksheet.write(rownum+16,1,col[14])
    worksheet.write(rownum+17,1,col[15])
    worksheet.write(rownum+18,1,col[16])
    worksheet.write(rownum+19,1,col[17])
    worksheet.write(rownum+20,1,col[18]) 
    worksheet.write(rownum+21,1,col[19]) 
    worksheet.write(rownum+22,1,col[20])
    worksheet.write(rownum+23,1,col[21])
    worksheet.write(rownum+24,1,col[22])
    worksheet.write(rownum+25,1,col[23])
    worksheet.write(rownum+26,1,col[24])
    worksheet.write(rownum+27,1,col[25])
    worksheet.write(rownum+28,1,col[26])
    worksheet.write(rownum+29,1,col[27])
    worksheet.write(rownum+30,1,col[28]) 
    
    
    #MAIN Renewal numbers
    worksheet.write(rownum+1,2,col[29]) 
    worksheet.write(rownum+2,2,col[30])
    worksheet.write(rownum+3,2,col[31])
    worksheet.write(rownum+4,2,col[32])
    worksheet.write(rownum+5,2,col[33])
    worksheet.write(rownum+6,2,col[34])
    worksheet.write(rownum+7,2,col[35])
    worksheet.write(rownum+9,2,col[36])
    worksheet.write(rownum+10,2,col[37])
    worksheet.write(rownum+11,2,col[38]) 
    worksheet.write(rownum+12,2,col[39])
    worksheet.write(rownum+13,2,col[40])
    worksheet.write(rownum+14,2,col[41])
    worksheet.write(rownum+15,2,col[42])
    worksheet.write(rownum+16,2,col[43])
    worksheet.write(rownum+17,2,col[44])
    worksheet.write(rownum+18,2,col[45])
    worksheet.write(rownum+19,2,col[46])
    worksheet.write(rownum+20,2,col[47])
    worksheet.write(rownum+21,2,col[48]) 
    worksheet.write(rownum+22,2,col[49])
    worksheet.write(rownum+23,2,col[50])
    worksheet.write(rownum+24,2,col[51])
    worksheet.write(rownum+25,2,col[52])
    worksheet.write(rownum+26,2,col[53])
    worksheet.write(rownum+27,2,col[54])
    worksheet.write(rownum+28,2,col[55])
    worksheet.write(rownum+29,2,col[56])
    worksheet.write(rownum+30,2,col[57])

   
   

worksheet = workbook.add_worksheet('Byram')



eformat= workbook.add_format({'text_wrap': True, 'valign': 'top' , 'num_format': 'mm/dd/yy'})
eformatlabel= workbook.add_format({'text_wrap': False, 'valign': 'top', 'bold': True})
    
worksheet.set_header('Byram')

worksheet.set_column(0,0,29)
worksheet.set_column(1,1,10)
worksheet.set_column(2,2,10)



worksheet.write(0,1,'Checkouts', eformatlabel)
worksheet.write(0,2,'Renewals', eformatlabel)
worksheet.write(1,0,'0 Adult Books', eformatlabel)
worksheet.write(2,0,'2 Adult Music', eformatlabel)
worksheet.write(3,0,'3 Juv Music', eformatlabel)
worksheet.write(4,0,'6 REF Main', eformatlabel)
worksheet.write(5,0,'7 Lending Art', eformatlabel)
worksheet.write(6,0,'8 ILL', eformatlabel)
worksheet.write(7,0,'9 Periodicals', eformatlabel)
worksheet.write(8,0,'10 Perrot Tote Bag Books', eformatlabel)
worksheet.write(9,0,'16 JUV Spknwrd Cass.', eformatlabel)
worksheet.write(10,0,'17 JUV music Cass.', eformatlabel)
worksheet.write(11,0,'18 Park/Museum Pass', eformatlabel)
worksheet.write(12,0,'20 Juvenile Books', eformatlabel)
worksheet.write(13,0,'22 Adult Spknwrd CDs', eformatlabel)
worksheet.write(14,0,'23 Juv Spnwrd CDs', eformatlabel)
worksheet.write(15,0,'31 Perr-A Express Books', eformatlabel)
worksheet.write(16,0,'37 Perr-J Periodicals', eformatlabel)
worksheet.write(17,0,'38 Perr-A Playaways', eformatlabel)
worksheet.write(18,0,'39 Perr-J Playaways', eformatlabel)
worksheet.write(19,0,'40 E-readers', eformatlabel)
worksheet.write(20,0,'41 Adult Non-Fict DVDs', eformatlabel)
worksheet.write(21,0,'42 DVD Children', eformatlabel)
worksheet.write(22,0,'44 DVD Adult', eformatlabel)
worksheet.write(23,0,'46 Perr-A Express DVDs', eformatlabel)
worksheet.write(24,0,'50 E-Books', eformatlabel)
worksheet.write(25,0,'60 J Books with CD ', eformatlabel)
worksheet.write(26,0,'61 Juv Multimedia Kit', eformatlabel)
worksheet.write(27,0,'62 Adult Game', eformatlabel)
worksheet.write(28,0,'63 Juv Game', eformatlabel)
worksheet.write(29,0,'80 Audbk Download', eformatlabel)



for rownum, col in enumerate(rows):
 #BYR Checkout numbers
    worksheet.write(rownum+1,1,col[66])
    worksheet.write(rownum+2,1,col[67])
    worksheet.write(rownum+3,1,col[68])
    worksheet.write(rownum+4,1,col[69])
    worksheet.write(rownum+5,1,col[70])
    worksheet.write(rownum+6,1,col[71])
    worksheet.write(rownum+7,1,col[72])
    worksheet.write(rownum+8,1,col[73])
    worksheet.write(rownum+9,1,col[74]) 
    worksheet.write(rownum+10,1,col[75]) 
    worksheet.write(rownum+11,1,col[76])
    worksheet.write(rownum+12,1,col[77]) 
    worksheet.write(rownum+13,1,col[78])     
    worksheet.write(rownum+14,1,col[79])    
    worksheet.write(rownum+15,1,col[80])
    worksheet.write(rownum+16,1,col[81])
    worksheet.write(rownum+17,1,col[82])
    worksheet.write(rownum+18,1,col[83])
    worksheet.write(rownum+19,1,col[84]) 
    worksheet.write(rownum+20,1,col[85]) 
    worksheet.write(rownum+21,1,col[86])
    worksheet.write(rownum+22,1,col[87])
    worksheet.write(rownum+23,1,col[88])
    worksheet.write(rownum+24,1,col[89])
    worksheet.write(rownum+25,1,col[90])
    worksheet.write(rownum+26,1,col[91])
    worksheet.write(rownum+27,1,col[92])
    worksheet.write(rownum+28,1,col[93])
    worksheet.write(rownum+29,1,col[94]) 
    #MAIN Renewal numbers
    worksheet.write(rownum+1,2,col[95]) 
    worksheet.write(rownum+2,2,col[96])
    worksheet.write(rownum+3,2,col[97])
    worksheet.write(rownum+4,2,col[98])
    worksheet.write(rownum+5,2,col[99])
    worksheet.write(rownum+6,2,col[100])
    worksheet.write(rownum+7,2,col[101])
    worksheet.write(rownum+8,2,col[102])
    worksheet.write(rownum+9,2,col[103])
    worksheet.write(rownum+10,2,col[104]) 
    worksheet.write(rownum+11,2,col[105])
    worksheet.write(rownum+12,2,col[106])
    worksheet.write(rownum+13,2,col[107])
    worksheet.write(rownum+14,2,col[108])
    worksheet.write(rownum+15,2,col[109])
    worksheet.write(rownum+16,2,col[110])
    worksheet.write(rownum+17,2,col[111])
    worksheet.write(rownum+18,2,col[112])
    worksheet.write(rownum+19,2,col[113])
    worksheet.write(rownum+20,2,col[114]) 
    worksheet.write(rownum+21,2,col[115])
    worksheet.write(rownum+22,2,col[116])
    worksheet.write(rownum+23,2,col[117])
    worksheet.write(rownum+24,2,col[118])
    worksheet.write(rownum+25,2,col[119])
    worksheet.write(rownum+26,2,col[120])
    worksheet.write(rownum+27,2,col[121])
    worksheet.write(rownum+28,2,col[122])
    worksheet.write(rownum+29,2,col[123])

  
    

worksheet.set_landscape()
worksheet.hide_gridlines(0)


worksheet = workbook.add_worksheet('Cos Cob')


worksheet.set_landscape()
worksheet.hide_gridlines(0)


eformat= workbook.add_format({'text_wrap': True, 'valign': 'top' , 'num_format': 'mm/dd/yy'})
eformatlabel= workbook.add_format({'text_wrap': False, 'valign': 'top', 'bold': True})
    
worksheet.set_header('Cos Cob')



worksheet.set_column(0,0,29)
worksheet.set_column(1,1,10)
worksheet.set_column(2,2,10)

worksheet.write(0,1,'Checkouts', eformatlabel)
worksheet.write(0,2,'Renewals', eformatlabel)
worksheet.write(1,0,'0 Adult Books', eformatlabel)
worksheet.write(2,0,'2 Adult Music', eformatlabel)
worksheet.write(3,0,'3 Juv Music', eformatlabel)
worksheet.write(4,0,'6 REF Main', eformatlabel)
worksheet.write(5,0,'7 Lending Art', eformatlabel)
worksheet.write(6,0,'8 ILL', eformatlabel)
worksheet.write(7,0,'9 Periodicals', eformatlabel)
worksheet.write(8,0,'10 Perrot Tote Bag Books', eformatlabel)
worksheet.write(9,0,'16 JUV Spknwrd Cass.', eformatlabel)
worksheet.write(10,0,'17 JUV music cass.', eformatlabel)
worksheet.write(11,0,'18 Park/Museum Pass', eformatlabel)
worksheet.write(12,0,'20 Juvenile Books', eformatlabel)
worksheet.write(13,0,'22 Adult Spknwrd CDs', eformatlabel)
worksheet.write(14,0,'23 Juv Spnwrd CDs', eformatlabel)
worksheet.write(15,0,'31 Perr-A Express Books', eformatlabel)
worksheet.write(16,0,'37 Perr-J Periodicals', eformatlabel)
worksheet.write(17,0,'38 Perr-A Playaways', eformatlabel)
worksheet.write(18,0,'39 Perr-J Playaways', eformatlabel)
worksheet.write(19,0,'40 E-readers', eformatlabel)
worksheet.write(20,0,'41 Adult Non-Fict DVDs', eformatlabel)
worksheet.write(21,0,'42 DVD Children', eformatlabel)
worksheet.write(22,0,'44 DVD Adult', eformatlabel)
worksheet.write(23,0,'46 Perr-A Express DVDs', eformatlabel)
worksheet.write(24,0,'50 E-Books', eformatlabel)
worksheet.write(25,0,'60 J Books with CD ', eformatlabel)
worksheet.write(26,0,'61 Juv Multimedia Kit', eformatlabel)
worksheet.write(27,0,'62 Adult Game', eformatlabel)
worksheet.write(28,0,'63 Juv Game', eformatlabel)
worksheet.write(29,0,'80 Audbk Download', eformatlabel)

for rownum, col in enumerate(rows):
    worksheet.write(rownum+1,1,col[130])
    worksheet.write(rownum+2,1,col[131])
    worksheet.write(rownum+3,1,col[132])
    worksheet.write(rownum+4,1,col[133])
    worksheet.write(rownum+5,1,col[134])
    worksheet.write(rownum+6,1,col[135])
    worksheet.write(rownum+7,1,col[136])
    worksheet.write(rownum+8,1,col[137])
    worksheet.write(rownum+9,1,col[138]) 
    worksheet.write(rownum+10,1,col[139]) 
    worksheet.write(rownum+11,1,col[140])
    worksheet.write(rownum+12,1,col[141]) 
    worksheet.write(rownum+13,1,col[142])     
    worksheet.write(rownum+14,1,col[143])    
    worksheet.write(rownum+15,1,col[144])
    worksheet.write(rownum+16,1,col[145])
    worksheet.write(rownum+17,1,col[146])
    worksheet.write(rownum+18,1,col[147])
    worksheet.write(rownum+19,1,col[148]) 
    worksheet.write(rownum+20,1,col[149]) 
    worksheet.write(rownum+21,1,col[150])
    worksheet.write(rownum+22,1,col[151])
    worksheet.write(rownum+23,1,col[152])
    worksheet.write(rownum+24,1,col[153])
    worksheet.write(rownum+25,1,col[154])
    worksheet.write(rownum+26,1,col[155])
    worksheet.write(rownum+27,1,col[156])
    worksheet.write(rownum+28,1,col[157])
    worksheet.write(rownum+29,1,col[158]) 
    #MAIN Renewal number
    worksheet.write(rownum+1,2,col[159]) 
    worksheet.write(rownum+2,2,col[160])
    worksheet.write(rownum+3,2,col[161])
    worksheet.write(rownum+4,2,col[162])
    worksheet.write(rownum+5,2,col[163])
    worksheet.write(rownum+6,2,col[164])
    worksheet.write(rownum+7,2,col[165])
    worksheet.write(rownum+8,2,col[166])
    worksheet.write(rownum+9,2,col[167])
    worksheet.write(rownum+10,2,col[168]) 
    worksheet.write(rownum+11,2,col[169])
    worksheet.write(rownum+12,2,col[170])
    worksheet.write(rownum+13,2,col[171])
    worksheet.write(rownum+14,2,col[172])
    worksheet.write(rownum+15,2,col[173])
    worksheet.write(rownum+16,2,col[174])
    worksheet.write(rownum+17,2,col[175])
    worksheet.write(rownum+18,2,col[176])
    worksheet.write(rownum+19,2,col[177])
    worksheet.write(rownum+20,2,col[178]) 
    worksheet.write(rownum+21,2,col[179])
    worksheet.write(rownum+22,2,col[180])
    worksheet.write(rownum+23,2,col[181])
    worksheet.write(rownum+24,2,col[182])
    worksheet.write(rownum+25,2,col[183])
    worksheet.write(rownum+26,2,col[184])
    worksheet.write(rownum+27,2,col[185])
    worksheet.write(rownum+28,2,col[186])
    worksheet.write(rownum+29,2,col[187])

    
workbook.close()


#Creating the email message
msg = MIMEMultipart()
msg['From'] = emailfrom
if type(emailto) is list:
    msg['To'] = ','.join(emailto)
else:
    msg['To'] = emailto
msg['Date'] = formatdate(localtime = True)
msg['Subject'] = emailsubject
msg.attach (MIMEText(emailmessage))
part = MIMEBase('application', "octet-stream")
part.set_payload(open(excelfile,"rb").read())
encoders.encode_base64(part)
part.add_header('Content-Disposition','attachment; filename=%s' % excelfile)
msg.attach(part)

#Sending the email message
smtp = smtplib.SMTP(emailhost, emailport)
smtp.sendmail(emailfrom, emailto, msg.as_string())
smtp.quit()








