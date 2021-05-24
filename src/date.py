from datetime import datetime
datetime_str = '25/12/20 11:12:13'

# Convert String ( ‘DD/MM/YY HH:MM:SS ‘) to datetime object
datetime_obj = datetime.strptime(datetime_str, '%d/%m/%y %H:%M:%S')

print(datetime_obj)
print('Type of the object:')
print(type(datetime_obj))