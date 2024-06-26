import requests
import operator
import pandas as pd

# Define the URL of the JSON file
url = "https://data.ny.gov/api/views/5xaw-6ayf/rows.json?accessType=DOWNLOAD"

def fetch_json_data(url):
    try:
        response = requests.get(url)
        response.raise_for_status()  # Check if the request was successful
        return response.json()  # Parse and return the JSON data
    except requests.exceptions.HTTPError as http_err:
        print(f"HTTP error occurred: {http_err}")
    except Exception as err:
        print(f"An error occurred: {err}")
        return None

def process_data(data):
    if 'data' in data:
        main_data = data['data']
    else:
        print("Main data section not found.")
        return None

    extracted_data = []

    for entry in main_data:
        if len(entry) > 10:  # Ensure there are enough elements in the entry
            extracted_entry = {
                "mega_mill_numbers": entry[9],  # Line 9
                "mega_ball": entry[10]  # Line 10
            }
            extracted_data.append(extracted_entry)
        else:
            print("Entry does not have enough elements:", entry)

    df = pd.DataFrame(extracted_data)
    return df



# will get lottery numbers in a list
def get_lottery():
    # Now, read the CSV and process the 'mega_mill_numbers'
    csv_path = r"lotto_data.csv"

    # Load the CSV file into a DataFrame
    df = pd.read_csv(csv_path)

    # Split the 'mega_mill_numbers' column into individual numbers and flatten the list
    mega_mill_numbers_list = df['mega_mill_numbers'].apply(lambda x: x.split())
    mega_mill_numbers_list = [int(number) for sublist in mega_mill_numbers_list for number in sublist]


    # Print the list to verify
    return mega_mill_numbers_list

def get_mega_numbers():
    # Now, read the CSV and process the 'mega_ball'
    csv_path = r"lotto_data.csv"

    # Load the CSV file into a DataFrame
    df = pd.read_csv(csv_path)

    # Ensure all values in 'mega_ball' column are converted to strings
    df['mega_ball'] = df['mega_ball'].astype(str)

    # Split the 'mega_ball' column into individual numbers and flatten the list
    def split_and_convert(x):
        try:
            return [int(num) for num in x.split()]
        except AttributeError:
            return []  # Handle non-string values gracefully

    mega_ball_list = df['mega_ball'].apply(split_and_convert)

    # Print the list to verify
    return mega_ball_list

# create a dictionary off drawn non mega numbers
def lottery_dict_counts(five):
    mega_dict = {}
    for numbers in range(len(five)):
        num1 = five[numbers]
        if num1 not in mega_dict.keys():
            mega_dict[num1] = 1
        else:
            mega_dict[num1] += 1

    return mega_dict

#create a dictionary off drawn mega numbers
def lottery_megaball_dict_counts(mega_ball):
    mega_ball_dict = {}
    for sublist in mega_ball:
        for num1 in sublist:
            if num1 not in mega_ball_dict:
                mega_ball_dict[num1] = 1
            else:
                mega_ball_dict[num1] += 1

    return mega_ball_dict

def check_value(num_dict, user_input):
    # Check if the user input is in the dictionary
    if user_input in num_dict:
        # Return the value associated with the input key
        return num_dict[user_input]
    else:
        # Return None if the input is not valid
        return None


def highest_draws(all_five, mega_ball):
    sorted_items = sorted(all_five.items(), key=operator.itemgetter(1), reverse=True)
    top_5 = sorted_items[:5]

    sorted_items = sorted(mega_ball.items(), key=operator.itemgetter(1), reverse=True)
    mega_top5 = sorted_items[:5]

    print("Top 5 most frequently drawn numbers:")
    for key, value in top_5:
        print(f"Number {key}: Drawn {value} times")

    print("\nTop 5 most frequently drawn Mega numbers:")
    for key, value in mega_top5:
        print(f"Number {key}: Drawn {value} times")


def user_selection(all_five):
    print("Enter five numbers between 1 and 70:")
    user_inputs = []
    for _ in range(5):
        user_input = int(input(f"Enter number {_ + 1}: "))
        user_inputs.append(user_input)

    for user_input in user_inputs:
        value = check_value(all_five, user_input)
        if value is not None:
            print(f"The number {user_input} was drawn: {value} times")
        else:
            print(f"Invalid number: {user_input}. Please enter a number between 1 and 70.")


def user_megaball_selection(mega_ball):
    while True:
        try:
            user_input = int(input("\nEnter a Mega number between 1 and 25: "))
            if user_input in mega_ball:
                print(f"The number {user_input} was drawn: {mega_ball[user_input]} times")
                break
            else:
                print(f"Invalid number: {user_input}. Please enter a number between 1 and 25.")
        except ValueError:
            print("Invalid input. Please enter a valid number.")




def main():
    # Store the list into a variable
    mega_numbers = get_lottery()
    # Store the mega ball list into a variable
    mega_ball_numbers = get_mega_numbers()

    # Create a dictionary of counts
    all_five = lottery_dict_counts(mega_numbers)
    mega_ball = lottery_megaball_dict_counts(mega_ball_numbers)


    while True:
        user_selection(all_five)
        user_megaball_selection(mega_ball)
        user_top_five = input(f"\nWould you like to see top 5 picked Mega Millions numbers? Y/N: ")
        if user_top_five == "y":
            highest_draws(all_five, mega_ball)
            break
        else:

            break



if __name__ == "__main__":

    # Fetch the JSON data from the URL
    data = fetch_json_data(url)

    if data:
        df = process_data(data)
        if df is not None:
            df.to_csv(r"W:\Pycharm Project\Lotto Bot\lotto_data.csv", index=False)
        else:
            print("Data processing failed.")
    else:
        print("Failed to retrieve data")

    main()