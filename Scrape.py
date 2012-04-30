import urllib, urllib2
from mmap import mmap,ACCESS_READ
import xlrd, xlwt
import xlutils
from xlutils.copy import copy

# Grab data from workbook
input_workbook = xlrd.open_workbook("C:\Users\Taizu\PyDev\\real-estate-scrape\Compliance_Report.xls")
input_sheet = input_workbook.sheet_by_index(0)

# Copy workbook from xlrd object to xlwt object and reads first sheet
write_book = copy(input_workbook)
write_sheet = write_book.get_sheet(0)
# write_sheet = (copy(xlrd.open_workbook("Longfei.xls"))).get_sheet(0)

# Define apt/pobox string filters
apt_list = [' apt', ' ste', ' suite', ' apartment', 'po ', 'p.o.', 'po.', 'box']
# Define engines
engines = {'zillow': {"link": "http://www.zillow.com/homedetails", "property_tag": "PropertyType", "property_offset": 16}, 'trulia': {'link': 'http://www.trulia.com/', 'property_tag': "Property type:", "property_offset": 10}}

# Loops through addresses in sheet and conditionally searches for address
# Writes to new excel file
for rownum in range(input_sheet.nrows-2):
	# returns address type result for zillow address query per address

	# Extract useful vars

	address_dict = {street_address = str(input_sheet.cell_value(rownum+2, 9)), street_number = street_address.split(' ')[0], street_name = street_address.split(' ')[1], city = str(input_sheet.cell_value(rownum+2, 11)), state =  str(input_sheet.cell_value(rownum+2, 12)), zip_code = str(input_sheet.cell_value(rownum+2, 13)[:5])}

	# Built generic search URL
	url_front = str("http://www.bing.com/search?q=" + address_dict[street_address].replace(' ', '+')) + '+' + address_dict[city].replace(' ', '+')) + '+' + address_dict[state].replace(' ', '+')) + '+' + address_dict[zip_code].replace(' ', '+')).replace('++', '+'))

	# Filter out apt/po searches
	if any(str(address_dict[street_address]).lower().find(word) != -1 for word in apt_list):
		write_sheet.write(rownum+2, 2, "Apartment/PO Box")

	# Core search: Try zillow, then trulia
	else:
		if search("zillow") != -1:
			write_sheet.write(rownum+2, 2, follow_link(result))
		elif search("trulia") != -1:
			write_sheet.write(rownum+2, 2, follow_link(result))
		else: write_sheet.write(rownum+2, 2, "Cannot find")

# ###############################################

# Save the book
write_book.save("C:\Users\Taizu\PyDev\\real-estate-scrape\Compliance_Report_10_Processed.xls")

# Utility functions
def search(engine):
	bing_search = urllib2.urlopen(url_front+"+"+engine).read()

	# If -1 or doesn't find required stuff (which you may or may not want to put in another function, depending)
	if bing_search.find(engines[engine]["link"]) == -1:
		return -1
	else:
		bing_start = bing_search.find(engines[engine]["link"])
		bing_end = bing_search[bing_start:].find('"')+bing_start
		engine_link = bing_search_zillow[bing_start:bing_end]
		if address_dict['street_number'] in engine_link and b in engine_link and c in engine_link and (d or e in engine_link):
			return engine_link
		else:
			return -1

def follow_link(result):
	# Open first Zillow link
	result = urllib2.urlopen(bing_search_zillow[bing_start:bing_end]).read()
	Property_Start = result.find(engines[engine]["property_type"], engines[engine]["property_offset"])
	Property_End = result[Property_Start:].find('"')+Property_Start
	property_type = result[Property_Start:Property_End]
	return property_type