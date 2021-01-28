import xlrd


xl_date = 44195
datetime_date = xlrd.xldate_as_datetime(xl_date, 0)
date_object = datetime_date.date()

print(date_object.year)
print(date_object.month)
print(date_object.day)


