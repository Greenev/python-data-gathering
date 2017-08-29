from bs4 import BeautifulSoup as BS
from excel_functions import cols
import re, requests, warnings

warnings.filterwarnings("ignore")

def get_html(url):
    resp = requests.get(url)
    return resp.text

def parse_by_inn(search_url, inn):
    soup = BS(get_html(search_url + inn), "html.parser")
    try:
        link = soup.find('div', class_ = 'grid').find('table').find('tbody').find('tr').find('td').find('a')
        link = re.split('\?+', link.get('href'))[0]
    except:
        link = ''
    return link

def parse_data(search_url, objects):
    for obj in objects:
        link = parse_by_inn(search_url, obj[cols[5]])
        soup = BS(get_html(re.split('/search', search_url)[0] + link), "html.parser")
        try:
            fr = soup.find('div', class_ = 'page').find('section', class_ = 'manager_info clearfix').find_all('div', class_ = 'fr')[1]
            rows = fr.find_all('p')
            obj[cols[1]] = re.split(':', rows[0].text)[1]
            obj[cols[2]] = re.split(':', rows[2].text)[1]
        except:
            obj[cols[1]] = ''
            obj[cols[2]] = ''
            
if __name__ == '__main__':
    print('This script provides html data parsing functions')
