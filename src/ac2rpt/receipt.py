import sys,os
import wx
from wx import xrc
import wx.grid as grd
import xlsxwriter
import re
from datetime import datetime
import time

#notascii = lambda s: len(s) != len(s.encode())
#def not_ascii(s):
#        return all(ord(c) > 128 for c in s)

#def is_rmb(s):
#    for c in s:
#        if ord(c) > 128:
#            print "ordc=",ord(c)
#        if ord(c) == 180:
#            c=u"\u00A5"
#            print "c=",c
#    return s
##        return all(ord(c) == 165 for c in s)

def export ( path, mapping, grid, payment_method, check_number ):
    """
        path: path to save the file
        mapping: mapping selected from mappings.py
        data: grid with csv data from report_utils.py
    """
    accounts={} # donor's account
    today = datetime.now().strftime('%Y%m%d')
    org_name=grid.GetOrgName()
    org_addr1=grid.GetOrgAddr1()
    org_addr2=grid.GetOrgAddr2()
    title=grid.GetTitle()
    period=grid.GetPeriod()
    isCR="No" 
    for row in range(grid.GetNumberRows()):
        # which account            
#        if mapping['skip'](row,grid): continue
        if row < 8: 
            continue
        
#        if 'Customer'|'Employee' in grid.grid_contents[row][2]: 
        
        a=['Customer','Employee','Vendor']
        if any(x in grid.grid_contents[row][2] for x in a) and grid.grid_contents[row][1].find("None") == -1: 
#        if any(x in grid.grid_contents[row][2] for x in a):
#        if 'Customer' in grid.grid_contents[row][2] or 'Employee' in grid.grid_contents[row][2] not grid.grid_contents[row][1].find("None"): 
            donor_name = grid.grid_contents[row][0]
            donor_id = grid.grid_contents[row][1]
            uacct="%s-%s" % (donor_name, donor_id)
            acct = accounts.setdefault(uacct,{})
            acct['ID'] = donor_id
            acct['Name'] = donor_name
            acct['TODAY'] = today
            acct['FileName']=donor_id + donor_name + '.xlsx'
#        uacct="%s-%s" % (mapping['BANKID'](row,grid), mapping['ACCTID'](row,grid))
         
        if grid.GetColNum(row)==9 and 'CR' in grid.grid_contents[row][2]:
            currency='USD'
            isCR="Yes" # check if this is still receipt or expense
#           print "Currency not the same."
            trans=acct.setdefault('trans',[])
            tran=dict([(k,mapping[k](row,grid)) for k in ['Trans. ID','Date','Memo','Amount Received','Amount Expended','Address','Account']]) 
            #        tran['TRNTYPE'] = tran['TRNAMT'] >0 and 'CREDIT' or 'DEBIT'
            acct['Address']=tran['Address']
            if re.search('zhao', tran['Memo'], re.IGNORECASE):
                tran['Memo']="For Zhaos' mission"
            elif re.search('globle', tran['Memo'], re.IGNORECASE):
                tran['Memo']="T mission in clouds"
            elif re.search('kun', tran['Memo'], re.IGNORECASE):
                tran['Memo']="For Ma Kun short mission"

            if grid.GetColNum(row+1)==9:
                acct['Address1']=grid.grid_contents[row+1][8]
#                if row+1<grid.GetNumberRows():
#                   if grid.GetColNum(row+2)==9:
#                       acct['Address2']=grid.grid_contents[row+2][8]
#                else:
#                    acct['Address2=']=' '
            else:
                acct['Address1']=' '
                acct['Address2']=' '
            if grid.GetColNum(row+2)==9:
                acct['Address2']=grid.grid_contents[row+2][8]
            else:
                acct['Address2']=' '
            print "Address2", acct['Name'],acct['Address2']
            
#            if not_ascii(tran['Amount Received']):
#               print "found non USD"

#            is_rmb(tran['Amount Received'])

#            if tran['Amount Received'].decode("utf8") > 128:
            method=payment_method.get(tran['Trans. ID']) 
            if method != None: 
               tran['Payment Method']=method
               tran['Check No.']=check_number[tran['Trans. ID']]
               tran['Currency']="USD"
            else: #  #couldn't find the transaction in bank register file "Make sure the periods are the same between Bank Register file and Card Transactions file. Transactions are not the same in these two files. No File exported."
               return 1

            if len(tran.get('Amount Received')) != 0:
               trans.append(tran)
        if grid.GetColNum(row)==8 and isCR=="Yes":
           total_received=grid.grid_contents[row][6].replace('$','').replace(',','') 
           total_expended=grid.grid_contents[row][7].replace('$','').replace(',','') 
           donor_total = float(total_received)-float(total_expended)
           acct['Total']='$' + str(donor_total)
           isCR="No"  # reset isCR
#    out=open(path,'w')


    for acct in accounts.values():
        if acct.get('Address') != None:
           path_file=os.path.join(path, acct['FileName'])
           workbook = xlsxwriter.Workbook(path_file)
           worksheet=workbook.add_worksheet()
           worksheet.set_portrait()
           worksheet.set_page_view()
           worksheet.center_horizontally()
           worksheet.set_paper(1)
           worksheet.set_margins(0.25,0.50,0.75,0.75)

           worksheet.set_column(0,1,10) # Trans. ID & Date, Account
           worksheet.set_column(2,2,25) #Memo
           worksheet.set_column(3,3,15) #Paymt. Method
           worksheet.set_column(4,4,10) #Check No.
           worksheet.set_column(5,5,15) #Amt. Received

           format_b = workbook.add_format({'bold':True})
           format_c = workbook.add_format({'align':'center'})
           format_bc = workbook.add_format({'bold':True,'align':'center'})
           format_bc14 = workbook.add_format({'bold':True,'align':'center','size':14})
#           format_bcw = workbook.add_format({'bold':True,'align':'center','size':11, 'pattern': 1, 'bg_color': 'blue', 'font_color': 'white'})
           format_bcw = workbook.add_format({'bold':True,'align':'center','size':11, 'top':1 ,'bottom':1})
#           format_bcw = workbook.add_format({'bold':True,'align':'center','size':11, 'pattern': 1, 'font_color': 'black'})
           format_bot_line = workbook.add_format({'top':1})

           worksheet.merge_range("A13:F13", 'Donation Receipt', format_bc14)
           worksheet.merge_range("A14:F14", period, format_bc14)
           worksheet.merge_range("A6:F6", org_name, format_bc14)

#           org_addr = org_addr1 + " " + org_addr2
#           worksheet.merge_range("A4:F4", org_addr, format_c)
           worksheet.write(1,2,"RECEIPT ENCLOSED",format_b)
           worksheet.write(0,0,"BFCI")
           worksheet.write(1,0,org_addr1)
           worksheet.write(2,0,org_addr2)

#           worksheet.insert_textbox('A9', period, options2)
           worksheet.write(7,5,"ID#:  %s" % acct['ID'],format_bc)
           worksheet.write(7,1,acct['Name'],format_b)
           worksheet.write(8,1,acct['Address'])
           worksheet.write(9,1,acct['Address1'])
#           if acct.get('Address2') != None:
#              worksheet.write(10,0,acct['Address2'])
       
#           for j in range(0,5):
#               cell = worksheet.table[14][j]
#               cell.format.set_top()

           worksheet.write(15,0,"Trans. ID",format_bcw)
           worksheet.write(15,1,"Date",format_bcw)
           worksheet.write(15,2,"Memo",format_bcw)
           worksheet.write(15,3,"Payment Method",format_bcw)
           worksheet.write(15,4,"Check No.",format_bcw)
           worksheet.write(15,5,"Amount Received",format_bcw)
           I=16
           for tran in acct['trans']:
               worksheet.write(I,0,tran['Trans. ID'],format_c)
               worksheet.write(I,1,tran['Date'],format_c)
               worksheet.write(I,2,tran['Memo'],format_c)
               worksheet.write(I,3,tran['Payment Method'],format_c)
               worksheet.write(I,4,tran['Check No.'],format_c)
               worksheet.write(I,5,tran['Amount Received'],format_c)
               I=I+1
        
           worksheet.merge_range("D%d:E%d"%(I+1,I+1), 'Deductable Amount:', format_bc)
           worksheet.write(I,5,acct['Total'],format_c)

           worksheet.merge_range("A%d:F%d"%(I+2,I+2), " ", format_bot_line)

           if re.search('zhao', tran['Memo'], re.IGNORECASE):
               worksheet.write(I+2,0,"If you would like to receive monthly update about Jian An & Laura Zhao's mission, please send an email to")
               worksheet.write(I+3,0,'love@missioninclouds.org with a subject as "Jian An & Laura Zhao"')
           worksheet.write(I+7,0,'our many thanks for your prayers and financial support.')
           range_string = "A15:F"+str(I+1)

           worksheet.write(I+5,0,'Thank you for your gift to the BFCI for the purpose of world missions! No goods or services were received for the')
           worksheet.write(I+6,0,'above contribution. Your Love for and confidence in the BFCI is greatly appreciated. We cannot adequately express')
           worksheet.write(I+7,0,'our many thanks for your prayers and financial support.')
           range_string = "A15:F"+str(I+1)
#           worksheet.write(39,0,'I thank my God Upon every remembrance of you, always in every prayer of mine for you all making request with joy.')
#           worksheet.write(40,0,'For your fellowship in the gospel from the first day until now;  Being confident of this very thing, that  he which hath')
#           worksheet.write(41,0,'begun a good work in you will perform it until the day of Jesus Christ.')
#    out.close()
#           apply_border_to_range(
#                workbook,
#                worksheet,
#                {
#                    "range_string": range_string,
#                    "border_style": 5,
#                },
#           )
           workbook.close()
    print "Exported %s" % path
    return 0
    
    
    
    

    
    
    
