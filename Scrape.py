import urllib, urllib2

from mmap import mmap,ACCESS_READ
import xlrd

# Collect addresses
addresses = []
book = xlrd.open_workbook("C:\Documents and Settings\LJiang\Desktop\Compliance Report.xls")
sheet=book.sheet_by_index(0)

# Loop through addresses in sheet and append to list
for rownum in range(sheet.nrows-2):
    addresses.append(str("http://www.bing.com/search?q=" + sheet.cell_value(rownum+2, 9).replace(' ', '+') + '+' + sheet.cell_value(rownum+2, 11).replace(' ', '+') + '+' + sheet.cell_value(rownum+2, 12).replace(' ', '+') + '+' + sheet.cell_value(rownum+2, 13)[:5]) + "+Zillow")

# Define single family list
single_family_addresses = []
row_counter = 0
# Loop through Bing-Zillow search
for address in range(len(addresses)):
	bing_result = urllib2.urlopen(addresses[address]).read()

	# Skip rest of iteration f can't find zillow on first page
	if bing_result.find("http://www.zillow.com/homedetails") == -1:
		continue

	bing_start = bing_result.find("http://www.zillow.com/homedetails")
	bing_end = bing_result[bing_start:].find('"')+bing_start

	# Open first Zillow link
	zillow_result = urllib2.urlopen(bing_result[bing_start:bing_end]).read()
	Property_Start = zillow_result.find('PropertyType", "')+16
	Property_End = zillow_result[Property_Start:].find('"')+Property_Start
	property_type = zillow_result[Property_Start:Property_End]
	print sheet.cell_value(row_counter + 2, 0)
	print property_type
	row_counter += 1