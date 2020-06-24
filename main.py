import xlrd
import datetime

path = 'corruption_vendor_dates.xlsx'

workbook = xlrd.open_workbook(path)
worksheet = workbook.sheet_by_index(0)

empty_cells = 0

rotation_1 = []
rotation_2 = []
rotation_3 = []
rotation_4 = []
rotation_5 = []
rotation_6 = []
rotation_7 = []
rotation_8 = []

corruption = {}
all_rotations = [rotation_1, rotation_2, rotation_3, rotation_4, rotation_5, rotation_6, rotation_7, rotation_8]
all_corruptions = []

start_datetime = datetime.datetime(year=2020, month=6, day=17, hour=8, minute=0, second=0)
current_datetime = datetime.datetime.now()

rotation_time = datetime.timedelta(days=3, hours=12)
end_datetime = start_datetime + rotation_time

start_dt_string = start_datetime.strftime("%d/%m/%Y %H:%M:%S")
current_dt_string = current_datetime.strftime("%d/%m/%Y %H:%M:%S")
end_dt_string = end_datetime.strftime("%d/%m/%Y %H:%M:%S")


def compare_dates(start, current, end):
    pass


current_list = []

# reads doc
for row in range(worksheet.nrows):
    for col in range(worksheet.ncols):
        value = worksheet.cell_value(row, col)
        # changes current active list
        if 'Rotation' in str(value):
            rotation_number = int(value[-1])
            # Sets current list as active list
            if rotation_number == 1:
                current_list = rotation_1
            if rotation_number == 2:
                current_list = rotation_2
            if rotation_number == 3:
                current_list = rotation_3
            if rotation_number == 4:
                current_list = rotation_4
            if rotation_number == 5:
                current_list = rotation_5
            if rotation_number == 6:
                current_list = rotation_6
            if rotation_number == 7:
                current_list = rotation_7
            if rotation_number == 8:
                current_list = rotation_8

            # print(f'------------------------------\nRotation number = {rotation_number}\n------------------------------')
            break
        # appends name rank value to dict objects
        if value != '':

            if col == 1:
                corruption['Name'] = str(value).strip()
            if col == 2:
                corruption['Rank'] = (int(value))
            if col == 3:
                corruption['Corruption Value'] = (int(value))
                # appends copy of current corruption to list of corruptions
                current_list.append(corruption.copy())
                # print(corruption)

            # print(f'row {row} col {col} {value}')

        else:
            empty_cells += 1

# appends dicts of corruptions to all_corruptions list
for curr_rotation in all_rotations:
    for corr in curr_rotation:
        all_corruptions.append(corr)

# sorts corruptions by value, name/rank/corruption value
all_corruptions = sorted(all_corruptions, key=lambda i: i['Name'])
all_corruptions = sorted(all_corruptions, key=lambda i: i['Rank'])
all_corruptions = sorted(all_corruptions, key=lambda i: i['Name'])


def bad_command():
    return print('Unaccepted Command')


def show_all_corruptions():
    last_used_name = ''
    print('---------------------------------------')
    for x in all_corruptions:
        if x['Name'] != last_used_name:
            print('\n')

        print(f'{x} ---------------- rotation: {get_rotation_number_from_corruption(x)} ---------------- Next available in: {check_next_rotation(x)}')
        last_used_name = x['Name']
    print('---------------------------------------')


def show_all_rotation_dates():
    datetime_check = start_datetime
    rotation_time = datetime.timedelta(days=3, hours=12)

    for x in range(len(all_rotations) * 3):
        start_dt_string = datetime_check.strftime("%d/%m/%Y %H:%M:%S")
        start_date = datetime.datetime.strptime(start_dt_string, "%d/%m/%Y %H:%M:%S")
        end_date = start_date + rotation_time
        end_dt_string = end_date.strftime("%d/%m/%Y %H:%M:%S")

        print(f'Rotation {x % 8 + 1} from: {start_dt_string} to: {end_dt_string}.')
        datetime_check = end_date

        if x % 8 + 1 == 8:
            print('------------------------------------------------------------------------------------')


def check_next_rotation(c):
    i = get_rotation_number_from_corruption(c) - 1
    datetime_check = start_datetime
    rotation_time = datetime.timedelta(days=3, hours=12)

    for x in range(len(all_rotations) * 3):
        start_dt_string = datetime_check.strftime("%d/%m/%Y %H:%M:%S")

        start_date = datetime.datetime.strptime(start_dt_string, "%d/%m/%Y %H:%M:%S")

        end_date = start_date + rotation_time

        end_dt_string = end_date.strftime("%d/%m/%Y %H:%M:%S")
        # print(f'>>>>>>>>>>Rotation {x + 1} from: {start_dt_string} to: {end_dt_string}.')
        datetime_check = end_date

        if i == x:
            # print(f'i == x : rotation {x+1} from: {start_dt_string} to: {end_dt_string}.')
            break
    next_one = datetime_check - current_datetime - rotation_time
    next_one -= datetime.timedelta(microseconds=next_one.microseconds)
    if next_one < datetime.timedelta(seconds=1):
        next_one += rotation_time * 8

    if next_one > datetime.timedelta(days=24, hours=12):
        next_one = 'Currently in rotation.'

    return next_one


def get_rotation_number_from_corruption(corruption):
    for current_rotation in all_rotations:
        # print(f'rotation number {all_rotations.index(current_rotation) + 1}')
        for corrupt in current_rotation:
            # print(corruption)
            # print(corruption['Name'])
            if corruption == corrupt:
                # print(f'{corruption} in rotation number {all_rotations.index(current_rotation) + 1}')
                return all_rotations.index(current_rotation) + 1


def get_rotation_number_from_date():
    datetime_check = start_datetime
    rotation_time = datetime.timedelta(days=3, hours=12)
    for x in range(len(all_rotations)):
        start_dt_string = datetime_check.strftime("%d/%m/%Y %H:%M:%S")
        # print("start date and time =", start_dt_string)

        start_date = datetime.datetime.strptime(start_dt_string, "%d/%m/%Y %H:%M:%S")

        end_date = start_date + rotation_time

        end_dt_string = end_date.strftime("%d/%m/%Y %H:%M:%S")

        next_rot = end_date - current_datetime
        next_rot -= datetime.timedelta(microseconds=next_rot.microseconds)

        # print("end date and time =", end_dt_string)
        # print(f'rotation{x+1} = {start_dt_string} - {end_dt_string}')

        if end_dt_string > current_dt_string:
            print(f'Currently on Rotation {x + 1} from: {start_dt_string} to: {end_dt_string}.')
            print(f'Next rotation in {next_rot}.')
            return x

        datetime_check = end_date


def show_current_rotation(num=0):
    print('----------------------------------------------------------------------------------------------------')
    rot_num = get_rotation_number_from_date()
    if rot_num > 7:
        rot_num -= 8
    current_rotation = all_rotations[rot_num + num]
    print('----------------------------------------------------------------------------------------------------')
    print(f'Rotation {rot_num + num + 1}')
    print('----------------------------------------------------------------------------------------------------')
    for x in current_rotation:
        print(f'{x}')
    print('----------------------------------------------------------------------------------------------------')
    return


def show_next_rotation():
    show_current_rotation(1)


def corruption_search(u_input):
    check = False
    last_used_name = ''
    for current_corruption in all_corruptions:
        if u_input in current_corruption['Name'].lower():
            # print(f'{current_corruption} in rotation number {get_rotation_number_from_corruption(current_corruption)}  '
            #       f'|||| Next available in: {check_next_rotation(current_corruption)}')

            if current_corruption['Name'] != last_used_name:
                print('------------------------------------\n')

            print(f'{current_corruption} ---------------- Next available in: {check_next_rotation(current_corruption)}')
            check = True
            last_used_name = current_corruption['Name']
    print('---------------------------------------')
    return check


accepted_commands = ['sort by name', 'sort by rank', 'sort by value', 'sort by corruption', 'quit']

hlp = '''------------------------------------------\n
Accepted commands are:

Quit - Quit the program
Dates - Shows all rotation dates
All - Shows all corruptions and ranks along with their corruption value
Next rotation - Shows next available corruptions
Current rotation - Shows currently available corruptions

Sort by X:
    Name
    Rank
    Corruption
    
Otherwise search for a corruption's name\n------------------------------------------'''


# show_current_rotation()
print('------------------------------------------\n|Type help to see all available commands.|\n------------------------------------------')
while True:
    user_input = input('\nWhat would you like to search for? ').lower().strip()

    if user_input == '':
        bad_command()
        continue

    if user_input == 'quit':
        break

    if user_input == 'help':
        print(hlp)
        continue

    if user_input == 'all':
        show_all_corruptions()
        continue

    if 'sort by' in str(user_input):
        split_input = user_input.split()
        if len(split_input) != 3:
            bad_command()
            continue
        else:
            if split_input[-1] == 'name':
                all_corruptions = sorted(all_corruptions, key=lambda i: i['Name'])
                print('Now Sorted by Name')
            elif split_input[-1] == 'rank':
                all_corruptions = sorted(all_corruptions, key=lambda i: i['Rank'])
                print('Now Sorted by Rank')
            elif split_input[-1] == 'corruption':
                all_corruptions = sorted(all_corruptions, key=lambda i: i['Corruption Value'])
                print('Now Sorted by Corruption')
            else:
                bad_command()
                continue

    if user_input == 'next':
        show_next_rotation()
        continue
    if user_input == 'current':
        show_current_rotation()
        continue
    if user_input == 'dates':
        show_all_rotation_dates()
        continue

    if not corruption_search(user_input):
        bad_command()



