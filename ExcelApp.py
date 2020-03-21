# Web Scraping and Excel Project:

from selenium import webdriver
import openpyxl
from datetime import datetime

class excelApp:
    def __init__(self):
        self.options = webdriver.ChromeOptions()

        # Makes browser invisible
        self.options.add_argument('headless')

        # Optional code for downloading PDFs:
        # self.options.add_experimental_option('prefs', {
        # "download.default_directory": r'C:\Users\Aaron\Desktop\VBA finance advisor app/pdf downloads', #Change default directory for downloads
        # "download.prompt_for_download": False, #To auto download the file
        # "download.directory_upgrade": True,
        # "plugins.always_open_pdf_externally": True #It will not show PDF directly in chrome
        # })

        self.linkToExcel()

        self.page = 1
        self.url = 'https://dash.lead.ac/repositories?page='+str(self.page)

        self.goToWebsite()
        
        # If not logged in:
        if self.driver.current_url == 'https://dash.lead.ac/users/sign_in':
            self.tryLoggingIn()

        # If still not logged in:
        if self.driver.current_url != self.url:
            self.loginError()
        
        self.excelTable = self.findExcelTable('Table1', self.sh._tables)

        self.getLastUpdate()
        
        # Counter determines whether Excel sheet is already up-to-date
        # If counter==0 when dates match, it is
        # If counter>0 when dates match, it isn't and stores new date
        self.counter = 0
        
        self.updateTables()
        
        self.saveAndExit()


    def getBrowser(self):
        self.driver = webdriver.Chrome(executable_path=r'C:\Users\Aaron\Desktop\VBA finance advisor app/chromedriver',options=self.options)
        
    def linkToExcel(self):
        self.wb = openpyxl.load_workbook(r'C:\Users\Aaron\Desktop\VBA finance advisor app\ExcelFile.xlsx')
        self.sh = self.wb['SheetName']

    def goToWebsite(self):
        self.driver.get(self.url)

    def tryLoggingIn(self):
        username = self.driver.find_element_by_id('user_email')
        password = self.driver.find_element_by_id('user_password')
        button = self.driver.find_element_by_name('button')

        username.send_keys('username')
        password.send_keys('password')
        button.click()

    def loginError(self):
        self.sh['errors'] = 'Login Error'
        self.driver.close()
        exit()

    @staticmethod
    def findExcelTable(table_name, tables):
        for table in tables:
            if table.displayName == table_name:
                return table
    
    def getLastUpdate(self):
        if type(self.sh['B1'].value)==datetime:
            self.last_updated = self.sh['B1'].value
        elif type(self.sh['B1'].value)==str:
            self.last_updated = datetime.strptime(self.sh['B1'].value,'%Y-%m-%d %H:%M%p')
        else:
            self.sh['errors'] = 'Last Updated cannot be {}'.format(type(self.sh['B1'].value))

    def updateTables(self):
        while True:
            if self.page != 1:
                self.goToWebsite()
            
            self.web_table = self.driver.find_elements_by_id('triage_form')
            self.web_rows = self.driver.find_elements_by_tag_name('tr')

            self.scrapeWebTable()

            self.page+=1

    def scrapeWebTable(self):
        for web_row in self.web_rows[1:]:
            tempRowData = []
            for col in web_row.find_elements_by_tag_name('td'):
                tempRowData.append(col.text)
            
            self.Returned_Date = datetime.strptime(tempRowData[7],'%Y-%m-%d %H:%M%p')
            if self.Returned_Date > self.last_updated:
                if self.counter == 0:
                    self.new_last_updated = self.Returned_Date
                
                self.First_name = tempRowData[2]
                self.Last_name = tempRowData[3]
                
                self.checkAndUpdateExcel()
                        
                # For additional features:            
                # Zip_code = tempRowData[4]
                # County = tempRowData[5]
            
                # PDF_link = web_row.find_element_by_link_text('PDF').get_attribute('href')
                
                self.counter += 1
            else:
                self.driver.close()
                if self.counter != 0:
                    self.sh['B1'] = self.new_last_updated

    def checkAndUpdateExcel(self):
        for excel_row in self.sh[self.excelTable.ref][1:]:
            if (excel_row[3].value == self.First_name) & (excel_row[4].value == self.Last_name):
                if excel_row[1].value==None:
                    excel_row[1].value=self.Returned_Date
                elif type(excel_row[1].value)==datetime:
                    if excel_row[1].value<self.Returned_Date:
                        pass
                        # TODO
                    elif excel_row[1].value>self.Returned_Date:
                        pass
                        # TODO
                else:
                    self.sh['errors']='The Received Date column in row {} must be empty or a date'.format(excel_row[0].row)
                
                break

    def saveAndExit(self):
        self.wb.save(r'C:\Users\Aaron\Desktop\VBA finance advisor app\ExcelFile.xlsx')
        exit()  

excelApp()
