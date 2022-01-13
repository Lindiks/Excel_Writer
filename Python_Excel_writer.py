import pandas as pd

xls = pd.read_excel('./headerTest.xlsx')

dfs = xls.to_dict()

is_all_day = False
is_all_day_input = input("Do you wish to add an all day event? [Y/N]: ")

if(is_all_day_input == 'Y'):
    is_all_day = True


for col_key in dfs:
    row_key = len(dfs[col_key])
    print(row_key)
    if(col_key == 'Subject'):  
        user_input = input('Event: ')
        dfs[col_key][row_key] = user_input

    if(col_key == 'Start Date'):  
        user_input = input('Start Date mm/dd/yyyy: ')
        dfs[col_key][row_key] = user_input

    if(col_key == 'End Date'):  
        user_input = input('End Date mm/dd/yyyy: ')
        dfs[col_key][row_key] = user_input

    if(col_key == 'Start Time' and not is_all_day):  
        user_input = input('Start Time tt:mm AM/PM: ')
        dfs[col_key][row_key] = user_input
        print(user_input)
    
    if(col_key == 'End Time' and not is_all_day ):  
        user_input = input('End Time tt:mm AM/PM: ')
        dfs[col_key][row_key] = user_input

    if(col_key == 'All Day Event' and is_all_day):  
        dfs[col_key][row_key] = True

    if(col_key == 'Description'):  
        user_input = input('Description: ')
        dfs[col_key][row_key] = user_input
    

dfs = pd.DataFrame(dfs)

dfs.to_excel('./headerTest.xlsx', index=False)










