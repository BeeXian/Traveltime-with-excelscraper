import simplejson, urllib.request, openpyxl


# x = ["118607","098585","648886","589318","538766","486038","464029"]


def getTravelTimes():
	# 1480321800
	x = scrapeExcelSheet()
	i = 0
	for i in range(len(x)):
		orig_coord = ["Singapore," + x[i]]
		dest_coord = ["Singapore,8,somapah,road"]
		url = "https://maps.googleapis.com/maps/api/distancematrix/json?units=imperial&origins={0}&destinations={1}&mode=transit&transit_mode=train&arrival_time=1480321800&key=AIzaSyAohbwRhNZ0UkgbnulEfip9cZYDlqMJBQM".format(str(orig_coord),str(dest_coord))
		result= simplejson.load(urllib.request.urlopen(url))
		# print(result)
		try:
			transit_time = result['rows'][0]['elements'][0]['duration']['text']
			# print(result)
		except KeyError as e:
			print("invalid address: " + str(x[i]))
			continue
		else:	
			print("Time taken: " + str(transit_time) + "   Address: " + str(x[i]))

	return "done"

def scrapeExcelSheet():

	holder = True

	while (holder == True):
		
		try:
			book = input("Enter excel book name: ")
			wb = openpyxl.load_workbook(book)
		except FileNotFoundError:
			print("couldn't find file: " + book )
			print("remember to include .xlsx")
			continue
		else:
			break


	while (holder == True):
		
		try:
			sheetnumber = input("Enter sheet name: ")
			sheet = wb.get_sheet_by_name(sheetnumber)
		except KeyError:
			print("Worksheet " + sheetnumber + " does not exist bro" )
			continue
		else:
			break


	# while (holder == True):
		
	# 	try:
	# 		columnNumber = int(input("Enter column number: "))
	# 		sheet.cell(row=1, column = columnNumber).value
	# 	except ValueError as e:
	# 		print(e)
	# 		continue
	# 	else:
	# 		break

	columnNumber = 3

	while (holder == True):
		
		try:
			start = int(input("Enter start row: "))
			sheet.cell(row=start, column = columnNumber).value

		except ValueError as e:
			print(e)
			continue
		else:
			break

	while (holder == True):
		
		try:
			end = int(input("Enter end row: ")) + 1
			sheet.cell(row=end, column = columnNumber).value
		except ValueError as e:
			print(e)
			continue
		else:
			break


	sheet = wb.get_sheet_by_name(sheetnumber)
	anslist = []

	print("\n" + "Extracting data from: " + book)
	print("Sheet: " + str(sheetnumber) + "\n" + "Column: " + str(columnNumber) + "  Row : " + str(start) +" to " + str(end - 1 ))

	for i in range(start,end):

		anslist.append(str(sheet.cell(row=i, column = columnNumber).value))
	
	print ("\n" + "Address list extracted: ")
	print (anslist)
	print()

	return(anslist)

# print(scrapeExcelSheet())
print(getTravelTimes())