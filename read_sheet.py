import gspread
import urllib.parse
from oauth2client.service_account import ServiceAccountCredentials
from search import search_using_automation
import threading

def initialize():
    # Specify the path to your service account JSON credentials file
    credentials_file = 'shining-courage-153713-1e879250b594.json'
    sheet_id = '1fqsHfl5ReI1LEsI3Gv_zxhmx7bxeVN2nNJKpSUPJgyQ'
    worksheet_id = '1424651226'

    # Define the scope for the Google Sheets API
    scope = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]

    # Authenticate with Google Sheets using your credentials
    creds = ServiceAccountCredentials.from_json_keyfile_name(credentials_file, scope)
    client = gspread.authorize(creds)

    # Open the Google Sheet by its title or URL
    sheet = client.open_by_key(sheet_id)  # Replace with your Google Sheet title or URL

    # Select the specific worksheet (by index or title)
    worksheet = sheet.get_worksheet_by_id(worksheet_id)  # Use 0 for the first worksheet, or use .title for a specific title
    return worksheet

def getNamedRangeValue(sheet, range_name):
    cell_values = sheet.get(range_name)
    return cell_values

# Function to get the named range index
def getNamedRangeIndex(sheet, range_name):
    cell_range = sheet.range(range_name)
    return (cell_range[0].row, cell_range[0].col)

def getInputs(sheet):
    # Function to get the named range value

    device = getNamedRangeValue(sheet, 'Device')[0][0]
    country = getNamedRangeValue(sheet, 'Country')[0][0]
    website = getNamedRangeValue(sheet, 'Website')[0][0]
    keywords = getNamedRangeValue(sheet, 'Keywords')
    rankingsStartPosition = getNamedRangeIndex(sheet, 'Ranking')
    websiteURLStartPosition = getNamedRangeIndex(sheet, 'WebsiteURL')
    keywordsStartPosition = getNamedRangeIndex(sheet, 'Keywords')

    return {
        'device': device,
        'country': country,
        'website': website,
        'keywords': keywords,
        'rankingsStartPosition': rankingsStartPosition,
        'websiteURLStartPosition': websiteURLStartPosition,
        'keywordsStartPosition': keywordsStartPosition
    }

def write_result_to_excel(worksheet, data):
    # data = [{link:123, rank=1}, {link:123, rank=1}]
    # iterate over data and create a 2D array of link
    link_array = []
    rank_array = []
    for value in data:
        sub_link_array = []
        sub_rank_array = []
        sub_link_array.append(value['link'])
        sub_rank_array.append(value['rank'])
        link_array.append(sub_link_array)
        rank_array.append(sub_rank_array)
       
    worksheet.update('Ranking', rank_array)
    worksheet.update('WebsiteURL', link_array)



def do_start(keywords, index, inputs, output):
    keyword = urllib.parse.quote(keywords)
    if keyword != '':
        result = search_using_automation(keyword, inputs['website'], inputs['device'])
        print(result)
        rank = '> 100' if result is None else result['rank']
        link = '' if result is None else result['link']
    
        output[index] = {
            'rank': rank, 
            'link': link
        }


def get_keyword_list(worksheet, keywordsStartPosition):
    keyword_list = ['pasta recipe']
    # for index in range(100):
        # keyword = worksheet.cell(keywordsStartPosition[0]+index, keywordsStartPosition[1]).value
        # if keyword is None or keyword == '':
            # break
        # keyword_list.append(keyword)
    return keyword_list 


def main():
    # Example usage
    worksheet = initialize()
    inputs = getInputs(worksheet)
    keywordsStartPosition = inputs['keywordsStartPosition']
    keyword_list = []
    keyword_list = get_keyword_list(worksheet, keywordsStartPosition)
    process = [None]*len(keyword_list)
    output = [None]*len(keyword_list)
    for index, keyword in enumerate(keyword_list):
        process[index] = threading.Thread(target= do_start, args=(keyword, index, inputs, output,))
        process[index].start()
    for i in range(len(process)):
        process[i].join()
    print(output)    
    write_result_to_excel(worksheet, output)

if __name__=="__main__":
    main()


    

