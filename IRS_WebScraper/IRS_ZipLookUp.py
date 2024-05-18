import requests
from bs4 import BeautifulSoup
import pandas as pd

#hard coded states
dic_states = {"alabama": 1,
              "alaska": 2,
              "american samoan:": 3,
              "arizona": 4,
              "arkansas": 5,
              "california": 6,
              "colorado": 7,
              "connecticut": 8,
              "delaware": 9,
              "district of columbia": 10,
              "florida": 11,
              "georgia": 12,
              "guam": 13,
              "hawaii": 14,
              "idaho": 15,
              "illinois": 16,
              "indiana": 17,
              "iowa": 18,
              "kansas": 19,
              "kentucky": 20,
              "louisiana": 21,
              "maine": 22,
              "maryland": 23,
              "massachusetts": 24,
              "michigan": 25,
              "minnesota": 26,
              "mississippi": 27,
              "missouri": 28,
              "montana": 29,
              "nebraska": 30,
              "nevada": 31,
              "new hampshire": 32,
              "new jersey": 33,
              "new mexico": 34,
              "new york": 35,
              "north carolina": 36,
              "north dakota": 37,
              "northern mariana": 38,
              "ohio": 39,
              "oklahoma": 40,
              "oregon": 41,
              "pennsylvania": 42,
              "puerto rico": 43,
              "rhode island": 44,
              "south carolina": 45,
              "south dakota": 46,
              "tennessee": 47,
              "texas": 48,
              "utah": 49,
              "vermont": 50,
              "virginia": 51,
              "virgin islands": 52,
              "washington": 53,
              "west virginia": 54,
              "wisconsin": 55,
              "wyoming": 56 }

#hard code column names for Pandas
info = ["Name of Business",
        "Address",
        "City/State/ZIP",
        "Point of Contact",
        "Telephone",
        "Type of Service"]

while True:
    #user input
    zipcode = input("Enter Zipcode: ")
    if len(zipcode)!=5:
        print(f"Invalid zipcode, please try again\n")
        continue
    state = input("Enter State: ")
    if state.lower() not in dic_states:
        print(f"Invalid State, please try again\n")
        continue
    #pages start at 0
    page = 0

    #States case insensitive
    get_state = dic_states.get(state.lower())

    #plug user input into url and request data
    url = f"https://www.irs.gov/efile-index-taxpayer-search?zip={zipcode}&state={get_state}&page={page}"
    response = requests.get(url)
    html = response.text

    #parse data setup
    soup = BeautifulSoup(html, 'html.parser')
    table = soup.find('table')
    providers = []

    #Iterate through pages and store data from tables into array providers
    next_page = soup.find("a", title="Go to next page")
    while next_page is not None:
        #tr holds rows of data in an html table, td are the columns in the html table
            for row in table.find_all('tr'):
                cells = row.find_all('td')
                #store all info into data array, then store provider information from data array into providers array
                data = [cell.text for cell in cells]
                providers.append(data[1])
            page += 1
            url = f"https://www.irs.gov/efile-index-taxpayer-search?zip={zipcode}&state={get_state}&page={page}"
            response = requests.get(url)
            html = response.text

            # parse data setup
            soup = BeautifulSoup(html, 'html.parser')
            table = soup.find('table')
            next_page = soup.find("a", title="Go to next page")

    #Split newlines and strip empty space
    d2 = [d.upper().splitlines() for d in providers]
    for space in range(len(d2)):
        d2[space].pop(-1)

    #Setup pandas table and import list
    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_columns', None)
    pd.set_option('display.expand_frame_repr', False)
    df = pd.DataFrame(d2, columns=info)

    search_option = input(f"\nPlease select a sorting option: \n"
                          f"1. Name of Business\n"
                          f"2. Address\n"
                          f"3. Point of Contact\n"
                          f"4. Telephone\n"
                          f"5. Type of Service\n")

    sort = {'1': "Name of Business",
            '2': "Address",
            '3': "Point of Contact",
            '4': "Telephone",
            '5': "Type of Service"}
    #default option if invalid input
    if search_option not in sort and search_option not in sort.values():
        search_option = '1'

    #output data
    if search_option in sort:
        print(df.sort_values(by=[sort.get(search_option)], ascending=True))
    elif search_option in sort.values():
        print(df.sort_values(by=[search_option], ascending=True))

    #ask user to quit or continue
    user_option = input("\nSearch again?\n")
    if user_option.lower() == "yes" or user_option.lower() == "y":
        continue
    break