from datetime import datetime,timedelta
test = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
test2 = (datetime.today() - timedelta(days=7)).strftime('%Y-%m-%d %H:%M:%S')
print(test)
print(test2)