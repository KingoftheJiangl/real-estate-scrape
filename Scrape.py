import urllib, urllib2
# from mmap import mmap,ACCESS_READ
import xlrd, xlwt, xlutils
from xlutils.copy import copy

# Grab data from workbook
input_workbook = xlrd.open_workbook("C:\Users\Taizu\PyDev\\real-estate-scrape\Compliance_Report.xls")
input_sheet = input_workbook.sheet_by_index(0)
# input_sheet = xlrd.open_workbook("C:\Users\Taizu\PyDev\\real-estate-scrape\Compliance_Report.xls").sheet_by_index(0)

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

	# Filter out apt/po searches
	if any(str(input_sheet.cell_value(rownum+2, 9)).lower().find(word) != -1 for word in apt_list):
		write_sheet.write(rownum+2, 2, "Apartment/PO Box")

	# Core search: Try zillow, then trulia
	elif:
		# loop while you have keys in engines dict
		if search("zillow") != -1:
			write_sheet.write(rownum+2, 2, follow_link(result))
	else: write_sheet.write(rownum+2, 2, "Check")

# ###############################################

# Save the book
write_book.save("C:\Users\Taizu\PyDev\\real-estate-scrape\Compliance_Report_10_Processed.xls")

# Utility functions
def search(engine):
	# Bing search pulling streetaddy, city, state, zip and interpolating +s and engine name
	bing_search = urllib2.urlopen(str("http://www.bing.com/search?q=" + str(input_sheet.cell_value(rownum+2, 9)).replace(' ', '+')) + '+' + input_sheet.cell_value(rownum+2, 11)).replace(' ', '+')) + '+' + str(input_sheet.cell_value(rownum+2, 12)).replace(' ', '+')) + '+' + str(input_sheet.cell_value(rownum+2, 13)[:5]).replace(' ', '+')).replace('++', '+'))+"+"+engine).read()

	# If -1 or doesn't find stuff (may want to move those conditionals up to this 'if' as well)
	if bing_search.find(engines[engine]["link"]) == -1:
		return -1
	else:
		bing_start = bing_search.find(engines[engine]["link"])
		bing_end = bing_search[bing_start:].find('"')+bing_start
		engine_link = bing_search_zillow[bing_start:bing_end]
		# TODO Add strong-count filter (<2 <strong>s) before the next </a>
		return engine_link

def follow_link(result):
	# Open first Zillow link
	result = urllib2.urlopen(bing_search_zillow[bing_start:bing_end]).read()
	Property_Start = result.find(engines[engine]["property_type"], engines[engine]["property_offset"])
	Property_End = result[Property_Start:].find('"')+Property_Start
	property_type = result[Property_Start:Property_End]
	return property_type