

from datetime import datetime
import csv

from wx.grid import PyGridTableBase


class Bank_Register_Grid(PyGridTableBase):
    """
        A very basic instance that allows the bank register of AccountEge contents to be used in a wx.Grid
        make sure the bank register report include the name of the
        organization(company), Company Address, Report Date and Time.
        This should be default, 
    """

    def __init__(self,csv_path,delimiter=',',skip_last=0):

        PyGridTableBase.__init__(self)
          # delimiter, quote could come from config file perhaps
        csv_reader = csv.reader(open(csv_path,'r'),delimiter=delimiter,quotechar='"')
        self.grid_contents = [row for row in csv_reader if len(row)>0]
        company_name = self.grid_contents[0]
        company_street = self.grid_contents[1]
        company_city = self.grid_contents[2]
        report_name = self.grid_contents[3] # Should be Bank Register 
        report_period = self.grid_contents[4] 
        report_date = self.grid_contents[5] 
        report_time = self.grid_contents[6] 

        self.grid_colnum = []
        self.payment_method={}
        self.check_number={}

        if skip_last:
            self.grid_contents=self.grid_contents[0:-skip_last]
        
        self.grid_rows = len(self.grid_contents)
        self.grid_cols = len(self.grid_contents[7])
        # the 8st row is the column headers
        for I in range (self.grid_rows):
           self.grid_colnum.append(len(self.grid_contents[I])) # setup the colnum of row I
           if len(self.grid_contents[I])>8: # check if there is column for payment method
               self.payment_method[self.grid_contents[I][1]] = self.grid_contents[I][8]
           elif len(self.grid_contents[I])>1:
               self.payment_method[self.grid_contents[I][1]] = ' '

           if len(self.grid_contents[I])>9: # check if there is column for payment method
               self.check_number[self.grid_contents[I][1]] = self.grid_contents[I][9]
           elif len(self.grid_contents[I])>1:
               self.check_number[self.grid_contents[I][1]] = ' '
        
        # header map
        # results in a dictionary of column labels to numeric column location            
        self.col_map=dict([(self.grid_contents[7][c],c) for c in range(self.grid_cols)])
        
    def GetNumberRows(self):
        return self.grid_rows
    
    def GetNumberCols(self):
        return self.grid_cols
    
    def GetColNum(self,row):
        return self.grid_colnum[row]

    def IsEmptyCell(self,row,col):
        if col<self.grid_colnum[row]:
           return len(self.grid_contents[row][col]) == 0
        else:
            return 0
    
    def GetValue(self,row,col):
        if col<self.grid_colnum[row]:
            return self.grid_contents[row][col]
        else:
            return ''
    
    def GetColLabelValue(self,col):
        return self.grid_contents[7][col]
    
    def GetColPos(self,col_name):
        return self.col_map[col_name]
    
def xmlize(dat):
    """
        Xml data can't contain &,<,>
        replace with &amp; &lt; &gt;
        Get newlines while we're at it.
    """
    return dat.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;').replace('\r\n',' ').replace('\n',' ')
    
def fromCSVCol(row,grid,col_name):
    """
        Uses the current row and the name of the column to look up the value from the csv data.
    """
    return xmlize(grid.GetValue(row,grid.GetColPos(col_name)))

    
class Card_Transaction_Grid(PyGridTableBase):
    """
        A very basic instance that allows the Card Transaction of AccountEge contents to be used in a wx.Grid
        make sure the Card Transaction report include the name of the
        organization(company), Company Address, Report Date and Time.
        This should be default, 
    """
    def __init__(self,csv_path,delimiter=',',skip_last=0):
    
        PyGridTableBase.__init__(self)
          # delimiter, quote could come from config file perhaps
        csv_reader = csv.reader(open(csv_path,'r'),delimiter=delimiter,quotechar='"')
        self.grid_colnum = []
        self.grid_contents = [row for row in csv_reader if len(row)>0]
        company_name = self.grid_contents[0]
        company_street = self.grid_contents[1]
        company_city = self.grid_contents[2]
        report_name = self.grid_contents[3] # Should be Bank Register 
        report_period = self.grid_contents[4] 
        report_date = self.grid_contents[5] 
        report_time = self.grid_contents[6] 


        if skip_last:
            self.grid_contents=self.grid_contents[0:-skip_last]
        
        self.grid_rows = len(self.grid_contents)
        self.grid_cols = len(self.grid_contents[7])
        # the 8st row is the column headers
        for I in range (self.grid_rows):
           self.grid_colnum.append(len(self.grid_contents[I])) # setup the colnum of row I

        # header map
        # results in a dictionary of column labels to numeric column location            
        self.col_map=dict([(self.grid_contents[7][c],c) for c in range(self.grid_cols)])
    def GetNumberRows(self):
        return self.grid_rows
    
    def GetNumberCols(self):
        return self.grid_cols
    
    def GetColNum(self,row):
        return self.grid_colnum[row]

    def IsEmptyCell(self,row,col):
        if col<self.grid_colnum[row]:
           return len(self.grid_contents[row][col]) == 0
        else:
            return 0
    
    def GetValue(self,row,col):
        if col<self.grid_colnum[row]:
            return self.grid_contents[row][col]
        else:
            return ''
    
    def GetColLabelValue(self,col):
        return self.grid_contents[7][col]
    
    def GetColPos(self,col_name):
        return self.col_map[col_name]
    
    def GetOrgName(self):
        return self.grid_contents[0][0]

    def GetOrgAddr1(self):
        return self.grid_contents[1][0]

    def GetOrgAddr2(self):
        return self.grid_contents[2][0]

    def GetPeriod(self):
        return self.grid_contents[4][0]

    def GetTitle(self):
        return self.grid_contents[3][0]

def fromCSVCol(row,grid,col_name):
    """
        Uses the current row and the name of the column to look up the value from the csv data.
    """
    return xmlize(grid.GetValue(row,grid.GetColPos(col_name)))

