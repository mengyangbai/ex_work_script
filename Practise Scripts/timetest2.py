
from dateutil import parser

def operate(object):
	if object is not None:
		object = parser.parse(object)
		object = object.strftime('%Y/%m/%d')
		return object	
	return object

dt = "2017-02-16 13:12"

print(operate(dt))