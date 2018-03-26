import sys, time
 
for i in range(5):
	print("123"+"de data")
	sys.stdout.write('~' * 10 + '\r')
	sys.stdout.flush()
	time.sleep(1)