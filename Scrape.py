import urllib, urllib2
from mmap import mmap,ACCESS_READ
import xlrd, xlwt
import xlutils
from xlutils.copy import copy

# zillow_search Procedure - Takes an address in form of [address, city, state, zip], bings the address for zillow and returns property type
# procedure url_filter checks zillow, trulia etc's url to ensure the list address and address that url refers to are one and the same(may need multiple)



# addresses - list collects addresses from excel sheet
addresses = []
input_workbook = xlrd.open_workbook("C:\Users\Taizu\PyDev\\real-estate-scrape\Compliance_Report_10 - Copy (4).xls")
input_sheet = input_workbook.sheet_by_index(0)
# uses xlutils to copy workbook from xlrd object to xlwt object and reads first sheet
write_book = copy(input_workbook)
write_sheet = write_book.get_sheet(0)
# Loops through addresses in sheet and appends to list object addresses
for rownum in range(input_sheet.nrows-2):
	# create sublist that contains address values: address1, city, state, zip
	sublist = [str(input_sheet.cell_value(rownum+2, 9)), str(input_sheet.cell_value(rownum+2, 11)), str(input_sheet.cell_value(rownum+2, 12)), str(input_sheet.cell_value(rownum+2, 13)[:5])]
	addresses.append(sublist)
print addresses


# address_types is a list that stores the address types in the list addresses after crawling the web
address_types = []

# Loop through addresses via bing for address types: zillow, trulia, homes.com
for address in addresses:
	# returns address type result for zillow address query per address
	if address[0].lower().find(" apt ") > -1 or address[0].lower().find(" ste ") > -1 or address[0].lower().find(" suite ") > -1: 
		address_types.append("Apartment")
	elif address[0].find("po ") > -1 or address[0].find(" box ") > -1:
		address_types.append("PO Box")

# TODO make this loop through zillow trulia and homes procedures as needed
	else:
		address_types.append(zillow_search(address))
	
print address_types



# Writes address_types list to correspondng rows in compliance report in the 3rd column.
for address_number, address in enumerate(input_sheet.col(0)):
	if not address_number:
		continue
	if address_number < 2:
		continue
	write_sheet.write(address_number,2,address_types[address_number-2])
	write_book.save("C:\Users\Taizu\PyDev\\real-estate-scrape\Compliance_Report_10 - Copy (4).xls")

def url_filter(verify_address, url_to_verify):
	print "lol"


def zillow_search(input_address_in_list_form):
	bing_search_zillow = urllib2.urlopen(str("http://www.bing.com/search?q=" +str(address[0].replace(' ', '+')) + '+' + str(address[1].replace(' ', '+')) + '+' + str(address[2].replace(' ', '+')) + '+' + str(address[3].replace(' ', '+')) + '+' + "+Zillow").replace('++', '+')).read()

	# if we don't find the beginning tag for zillow for the address, we move on to the next address. this will break when we introduce trulia and homes.com

	if bing_search_zillow.find("http://www.zillow.com/homedetails") == -1:
		property_type = "No address"
		return property_type
	elif bing_search_zillow.find("http://www.zillow.com/homedetails") != -1:
		# finds starting and ending positions for the zillow page on Bing search

		bing_start = bing_search_zillow.find("http://www.zillow.com/homedetails")
		bing_end = bing_search_zillow[bing_start:].find('"')+bing_start

		# Open first Zillow link
		zillow_result = urllib2.urlopen(bing_search_zillow[bing_start:bing_end]).read()
		Property_Start = zillow_result.find('PropertyType", "')+16
		Property_End = zillow_result[Property_Start:].find('"')+Property_Start
		property_type = zillow_result[Property_Start:Property_End]
		return property_type