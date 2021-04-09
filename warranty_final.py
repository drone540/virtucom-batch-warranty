from urllib.request import urlopen
from urllib.parse import quote
from datetime import date
import csv, re, argparse, os, json
os.system("")

parser = argparse.ArgumentParser()
parser.add_argument("-m", help="Include Model Information in output", action="store_true")
parser.add_argument("-p", help="Include Product Number in output", action="store_true")
parser.add_argument("-t", help="Include Quick Track Information in output", action="store_true")
parser.add_argument("-q", help="Sample Output Test Data", action="store_true")
parser.add_argument("-i", "--input", help="CSV file to process (default: warranty.csv)", default="warranty.csv", type=argparse.FileType('r'))
parser.add_argument("search", nargs='*', help="String to search for or column to search by")
args = parser.parse_args()
args.input.close()
if args.q:
	args.search = ['P2060DQL', 'NXG55AA0108171DA797600', '5CD911BWB9', 'PXXXXXXX', '967sgh2', 'V1907CA10170', '184348', 'arts', 'zz']
def getWarranty(get):
	if get[0].upper() == "V":
		page = urlopen('https://www.virtucom.com/eWAP/Entitlement/display.php?vciTagInput=' + get)
	elif get.isdecimal():
		page = urlopen('https://www.virtucom.com/eWAP/Entitlement/display.php?assetTagInput=' + get)
	else:
		page = urlopen('https://www.virtucom.com/eWAP/Entitlement/display.php?srnum=' + get)
	html_bytes = page.read()
	html = html_bytes.decode('utf-8')
	serialMatch = re.search('serialNum ?= ?"(.*)"', html)
	if serialMatch:
		serial = serialMatch.group(1)
		model =      re.search('model ?= ?"(.*)"', html).group(1)
		vci =  re.search('virtucomNum ?= ?"(.*)"', html).group(1)
		asset =   re.search('assetTag ?= ?"(.*)"', html).group(1)
		product=re.search('productNum ?= ?"(.*)"', html).group(1)
		tracker=re.search('tracker_status ?= ?"(.*)"', html).group(1)
		startDate = date.fromisoformat(re.search('Date\("(.+?)"', html).group(1))
		baseStan = int(re.search('baseStan ?= ?parseInt\("(.*)"', html).group(1))
		baseAdp =   int(re.search('baseAdp ?= ?parseInt\("(.*)"', html).group(1))
		baseBat =   int(re.search('baseBat ?= ?parseInt\("(.*)"', html).group(1))
		totalStan=int(re.search('totalStan ?= ?parseInt\("(.*)"', html).group(1))
		totalAdp = int(re.search('totalAdp ?= ?parseInt\("(.*)"', html).group(1))
		totalBat = int(re.search('totalBat ?= ?parseInt\("(.*)"', html).group(1))
		baseWarranty = startDate.replace(year=startDate.year + baseStan, day=startDate.day - 1)
		vciWarranty = startDate.replace(year=startDate.year + totalStan, day=startDate.day - 1)
		adpWarranty = startDate.replace(year=startDate.year + totalAdp, day=startDate.day - 1)
		batWarranty = startDate.replace(year=startDate.year + totalBat, day=startDate.day - 1)
		extension = re.search('extension_array ?= ?\[(.*)\]', html)
		today = date.today()
		warranty = ""
		if baseWarranty > today:
			warranty = "Base: " + baseWarranty.strftime("%x")
			if baseAdp > 0:
				warranty = warranty + " +ADP"
			if baseBat > 0:
				warranty = warranty + " +Battery"
		else:
			warranty = "Base: Expired"
		if vciWarranty > today:
			warranty = warranty + " | VCI: " + vciWarranty.strftime("%x")
			if totalAdp > 0 and adpWarranty > today:
				warranty = warranty + " +ADP"
			if totalBat > 0 and batWarranty > today:
				warranty = warranty + " +Battery"
		else:
			warranty = warranty + " | VCI: Expired"
		if extension.group(1):
			extended_info = json.loads(extension.group(1))
			extendedDate = date.fromisoformat(extended_info['ewd'])
			specialWarranty = extendedDate.replace(year=extendedDate.year + int(extended_info['ewy']), day=extendedDate.day - 1)
			if specialWarranty > today:
				warranty = warranty + " | Special: " + specialWarranty.strftime("%x") + " - Type: " + extended_info['ewt']
			else:
				warranty = warranty + " | Special: Expired"
	else:
		serial = "Not Found"
		vci = asset = warranty = product = model = tracker = "N/A"

	print("S:\033[93m"+ serial + "\033[0m | V:\033[96m" + vci + "\033[0m | A:\033[92m" + asset + "\033[0m | W:\033[94m" + warranty + "\033[0m")
	# print("S:\033[93m"+ serial + "\033[0m | V:\033[96m" + vci + "\033[0m | A:\033[92m" + asset + "\033[0m")
	# print("W:\033[94m" + warranty + "\033[0m")
	if args.t:
		print("T:\033[91m" + tracker + "\033[0m")
	if args.p:
		print("P:\033[32m" + product + "\033[0m")
	if args.m:
		print("M:\033[95m" + model + "\033[0m")
	return { 'serial': serial, 'vci': vci, 'asset': asset, 'warranty': warranty, 'product' : product, 'model' : model, 'tracker' : tracker }

for item in args.search:
	if item.isalpha() and len(item) <= 3:
		print("Searching using column letter " + item.upper())
		inputfile = args.input.name
		outputfile = os.path.splitext(inputfile)[0] + '_output.csv'
		excel_col_num = lambda a: 0 if a == '' else 1 + ord(a[-1]) - ord('A') + 26 * excel_col_num(a[:-1])
		column = excel_col_num(item.upper())-1
		with open(inputfile, 'r') as csvinput:
			row_count = sum(1 for row in csvinput)
		with open(inputfile, 'r') as csvinput:
			with open(outputfile, 'w') as csvoutput:
				reader = csv.reader(csvinput)
				writer = csv.writer(csvoutput, lineterminator='\n')
				for row in reader:
					if column > len(row)-1:
						parser.error('Column not found in CSV file: ' + inputfile)
					search = quote(row[column])
					print("Status: " + str(reader.line_num) + "/" + str(row_count))
					page = getWarranty(search)
					row.append(page['serial'])
					row.append(page['vci'])
					row.append(page['asset'])
					row.append(page['warranty'])
					if args.t:
						row.append(page['tracker'])
					if args.p:
						row.append(page['product'])
					if args.m:
						row.append(page['model'])
					writer.writerow(row)

	elif item.isalnum():
		print('Searching using string: ' + item)
		page = getWarranty(item)
	else:
		raise Exception('Invalid search input was entered. Please enter string or column letter')
