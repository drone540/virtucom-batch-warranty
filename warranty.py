inputfile = 'warranty.csv'
outputfile = 'warranty_output.csv'
set_column = 6

from urllib.request import urlopen
from urllib.parse import quote
import csv, re, urllib
column = set_column-1
with open(inputfile, 'r') as csvinput:
	row_count = sum(1 for row in csvinput)
	print(row_count)
with open(inputfile, 'r') as csvinput:
	with open(outputfile, 'w') as csvoutput:
		reader = csv.reader(csvinput)
		writer = csv.writer(csvoutput, lineterminator='\n')
		for row in reader:
			serialnum = quote(row[column])
			print("Status: " + str(reader.line_num) + "/" + str(row_count))
			page = urlopen('https://www.virtucom.com/eWAP/Entitlement/display.php?srnum=' + serialnum)
			html_bytes = page.read()
			html = html_bytes.decode("utf-8")
			vcimatch = re.search('var virtucomNum = "(.*)"', html)
			assetmatch = re.search('var assetTag= "(.*)"', html)
			if vcimatch:
				if not assetmatch.group(1):
					asset = "Asset N/A"
				else:
					asset = assetmatch.group(1)
				print("Serial: "+ serialnum + " | VCI: " + vcimatch.group(1) + " | Asset: " + asset)
				row.append(vcimatch.group(1))
				row.append(asset)
			else:
				print("Serial: "+ serialnum + " | VCI: " + "N/A" + " | Asset: " + "N/A")
				row.append("VCI N/A")
				row.append("Asset N/A")
			writer.writerow(row)
