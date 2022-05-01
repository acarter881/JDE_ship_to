import pdfplumber
import re
import pandas as pd
from tqdm import tqdm
from collections import namedtuple
 
class myShipToInfo:
    def __init__(self, pathToPDF) -> None:
        self.path = pathToPDF[1:-1] # Hold shift and right click the PDF file, clicking "Copy as path"
        self.invoiceNumberRegEx = re.compile(r'[78]\d{7}')
        self.rows = list()
        self.namedTuple = namedtuple('LIST', 'Ship_to_Name Invoice_Number City State Zip_Code')
 
    def main(self):
        with pdfplumber.open(path_or_fp=self.path) as pdf:
            for i in tqdm(range(len(pdf.pages))):
                self.page = pdf.pages[i]
                self.text = self.page.extract_text()
                self.invoiceNumber = re.search(pattern=self.invoiceNumberRegEx, string=self.text).group(0)
                
                # Crop the page to only display the section with the ship to information (i.e., name and address)
                self.shipToCrop = self.page.crop((340, 0.225 * float(self.page.height), self.page.width, self.page.height / 3))
                self.ship_to_name = self.shipToCrop.extract_text().split('\n')[1:][0]               
                
                # Take the last row of the ship to data (i.e., the city, state, zip)
                self.shipToData = self.shipToCrop.extract_text().split('\n')[-1]
                
                # Find the name of the city
                self.city = ' '.join(self.shipToData.split()[:-2])
                
                # The second to last item in the last row should be the two letter state abbreviation
                try:
                    self.state = self.shipToData.split()[-2]
                except IndexError:
                    print(f'There was an IndexError at page {self.page}')
                    self.state = 'Please review'
                
                # The last item in the last row should be the 5 or 9 digit zip code
                self.zipCode = self.shipToData.split()[-1]
 
                # If the document number is already in our list, move to the next page. Otherwise, write the relevant data to the "rows" list
                if re.search(rf'{self.invoiceNumber}', string=str(self.rows)):
                    pass 
                else:
                    self.rows.append(
                        (self.namedTuple(
                            self.ship_to_name, 
                            self.invoiceNumber, 
                            self.city, 
                            self.state, 
                            self.zipCode)
                        )
                                    )
        
    # Function to write data to Excel
    def toPandas(self):
        self.df = pd.DataFrame(data=self.rows)
        self.df.to_excel(excel_writer='For_JDE_Use_Tax_on_Samples.xlsx', 
                         sheet_name='Use Tax Ship Tos', 
                         index=False, 
                         freeze_panes=(1,0)
                        )
 
# Instantiate the class and run the necessary functions
if __name__ == '__main__':
    c = myShipToInfo(pathToPDF=input('Please input the path to your PDF file: '))
    c.main()
    c.toPandas()
