def compare_time(time1,time2):
    number1 = 0
    number2 = 0
    if time1 == "":
        number1 = 999999999
    else:
        row = time1.split("-")
        number1 = int(row[0])*365+int(row[1].lstrip("0"))*30+ int(row[2].lstrip("0"))
    if time2 == "":
        number2 = 999999999
    else:
        row = time2.split("-")
        number2 = int(row[0])*365+int(row[1].lstrip("0"))*30+ int(row[2].lstrip("0"))
    if number1==number2:
        return 0
    elif number1 > number2:
        return 1
    elif number1 < number2:
        return -1
        
print(compare_time("2019-01-01","2019-01-01"))