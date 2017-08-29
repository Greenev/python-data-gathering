#Script allows parsing of .xlsx source file to select required records (GKH companies)
#Then it finds contact info for each company by parsing https://www.reformagkh.ru
#New .xlsx file are forming with results was parsed

#Source file could be downloaded from http://78.mchs.gov.ru/document/4594587
#It should be resave as .xlsx before using with the script.

from excel_functions import cols, excel_parsing, db_create
from html_functions import parse_by_inn, parse_data

source_path = 'source_sample.xlsx'
target_path = 'test_db.xlsx'
search_url = 'https://www.reformagkh.ru/search/mo?query='

#source .xlsx file row numbers
start = 38
end = 230
      
def main():
    objects=[]
    #parsing source companies list from .xlsx
    excel_parsing(start, end, source_path, objects)
    #parsing companies contact info from 'search_url'
    parse_data(search_url, objects)
    #making new .xlsx database file
    db_create(target_path, objects)
    
if __name__ == '__main__':
    main()
