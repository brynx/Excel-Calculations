from xlutils.copy import *
from xlrd import *
from xlwt import *
from datetime import *
import ConfigParser

# CONFIG #
config = ConfigParser.ConfigParser()
config.read("config.txt")

# EXCEL #
rb = open_workbook(config.get("general", "filename"), formatting_info=True, on_demand=True)
r_sheet = rb.sheet_by_index(4)
wb = copy(rb)
w_sheet = wb.get_sheet(4)

# VARIABLES #
includedEvents = []
noBackInServiceTime = []
incidentTimes = []


time_value = xldate_as_tuple(r_sheet.cell(570,16).value,rb.datemode)
print(time_value)
time_value = time_value(3) + time_value(4)





# IF PM IS IN THE CELL_VALUE #
dateTest = r_sheet.cell_value(760,16)
print(dateTest)

dateTest = dateTest[:-3]

dateTest = dateTest[-5:].replace(" ", "").replace(":", "")
dateTest = int(dateTest)
dateTest = dateTest + 1200
print(dateTest)



def checkCellsIncluded():
	try:
		for x in range(572, 100000): # Maximum of 100,000 entries
			if r_sheet.cell_value(x, 27) == "Included Events":
				includedEvents.append(x)
	except IndexError:
		print(includedEvents)
		print("Scanned included events")

def removeEmpties():
	for i in noBackInServiceTime:
		if i in includedEvents:
			includedEvents.remove(i)

def checkCellsEmpty():
	for x in includedEvents:
		if r_sheet.cell_value(x, 26) == "":
				noBackInServiceTime.append(x)
	print(noBackInServiceTime)
	print("Scanned for empties")

def getTimes():
	for x in includedEvents:
		incidentTimes.append(r_sheet.cell_value(x, 16))
		#if "PM" in r_sheet.cell_value(x, 16):
		#	incidentTimes.append
	print(incidentTimes)
	print("Got times")


checkCellsIncluded()
checkCellsEmpty()
getTimes()


print("Number of items in that are included: " + str(len(includedEvents)))
removeEmpties()
print("Number of items included excluding incidents with empty 'Back in Service' times: " + str(len(includedEvents)))
print("Number of entries that have no 'Back in Service' entry: " + str(len(noBackInServiceTime)))


print(r_sheet.cell_type(570, 16))
