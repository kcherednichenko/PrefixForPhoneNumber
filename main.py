import pandas as pd
import logging


def is_full_range(prefix):
    global left, right
    current_left = int(prefix + '0' * (len(left) - len(prefix)))
    current_right = current_left + int('1' + '0' * (len(left) - len(prefix))) - 1
    if current_left < int(left) or current_right > int(right):
        return False
    return True


def is_partial_range(prefix):
    global left, right
    current_left = int(prefix + '0' * (len(left) - len(prefix)))
    current_right = current_left + int('1' + '0' * (len(left) - len(prefix))) - 1
    if current_right < int(left) or current_left > int(right):
        return False
    return True


def add_to_prefix_dict(prefix, zone):
    global prefix_dict
    if prefix in prefix_dict:
        prefix_dict[prefix].append(zone)
        zones = ', '.join(prefix_dict[prefix])
        logging.error(f"Duplication of prefix {prefix} in zones: {zones}")
    else:
        prefix_dict[prefix] = [zone]


def find_prefix(prefix, zone):
    global output_data
    if is_full_range(prefix):
        output_data.append([prefix, zone])
        add_to_prefix_dict(prefix, zone)
    else:
        for i in range(10):
            if is_full_range(prefix + str(i)):
                output_data.append([prefix + str(i), zone])
                add_to_prefix_dict(prefix + str(i), zone)
            elif is_partial_range(prefix + str(i)):
                find_prefix(prefix + str(i), zone)


logging.basicConfig(filename='info.log', filemode='w', format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                    level=logging.NOTSET)

print("Do you want to use default path and worksheet name of excel file? [y/n]")
answer = ''
path_of_file = 'input.xlsx'
sheet_name = 'Лист1'
while answer != 'y' and answer != 'n':
    answer = input()
    if answer == 'y':
        input_df = pd.read_excel(path_of_file, sheet_name=sheet_name)
    elif answer == 'n':
        print("Type space-separated path and worksheet name of excel file: ")
        path_of_file, sheet_name = input().split()
    else:
        print("Error! Type letter y or n in lowercase")

try:
    input_df = pd.read_excel(path_of_file, sheet_name=sheet_name)

except FileNotFoundError:
    logging.error(f"No such file or directory: {path_of_file}")
    print(f"Error! No such file or directory: {path_of_file}")

except ValueError:
    logging.error(f"Worksheet named {sheet_name} not found")
    print(f"Error! Worksheet named {sheet_name} not found")

else:
    input_data = []
    zones_dict = dict()
    prefix_dict = dict()
    for index, row in input_df.iterrows():
        start_range, finish_range = row['Диапазон'].split('-')
        start_range = str(row['Общий префикс']) + start_range
        finish_range = str(row['Общий префикс']) + finish_range
        zones_dict[row['Преф.зона']] = zones_dict.get(row['Преф.зона'], len(start_range))
        if zones_dict[row['Преф.зона']] == len(start_range) and zones_dict[row['Преф.зона']] == len(finish_range) \
                and int(start_range) <= int(finish_range):
            input_data.append([row['Преф.зона'], start_range, finish_range])
        else:
            logging.error(
                f"Incorrect ranges of phone numbers in zone: {row['Преф.зона']} in line {index + 2}: {start_range}, {finish_range}")
            print(f"Incorrect ranges of phone numbers: {start_range}, {finish_range}")
            exit()

    input_data.sort(key=lambda x: (x[0], int(x[1])))
    output_data = []
    i = 0
    while i < len(input_data):
        left = input_data[i][1]
        current_zone = input_data[i][0]
        right = ''
        while i + 1 < len(input_data) and input_data[i + 1][0] == current_zone and \
                int(input_data[i][2]) >= int(input_data[i + 1][1]):
            right = str(max(int(input_data[i][2]), int(input_data[i + 1][2])))  #
            i += 1
        if not right:
            right = input_data[i][2]

        j = 0
        prefix = ''
        while j < len(left) and left[j] == right[j]:
            prefix += left[j]
            j += 1

        find_prefix(prefix, current_zone)

        i += 1

    output_data.sort(key=lambda x: (int(x[1][5:]), int(x[0])))

    df = pd.DataFrame(output_data, columns=['Префикс', 'Преф.зона'])
    df.to_excel('output.xlsx', index=False)
    logging.info(f"Program finished successfully")
