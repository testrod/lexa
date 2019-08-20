#lexa script

import sys	#Library to provide access to some variables used
import xlrd #Library to extract data from Microsoft Excel

inputfile = sys.argv[1]		#inputfile name
outputfile = sys.argv[2]	#outputfile new_file
num_lines = 0				
total_num_lines = 0

workbook = xlrd.open_workbook(inputfile, on_demand = True) 	#open the xls file to be procesed
new_file = open(outputfile,"w")								#open a new txt file to save the results 
totalsheets = len(workbook.sheet_names())

print "This xls file has", totalsheets, "sheets"
print "The input file ", inputfile ," is being procesed"

i = 0
while i < totalsheets:

	sheet0 = workbook.sheet_by_index(i)						#set the first sheet to be procesed
	c = 2													#start from the second row
	rows = sheet0.nrows										#total of rows

	while c < rows:
		cl = sheet0.cell(c, 0).value
		clant = sheet0.cell(c-1, 0).value
		
		if (cl != '') :
			new_file.write(sheet0.cell(c, 0).value)
			new_file.write("\n")
			num_lines = num_lines+1
			
			
		c = c + 1
		
	i = i + 1
	
	print "sheet", i ,"finished"
	total_num_lines = total_num_lines + num_lines
	num_lines = 0
	
new_file.close()
print "The output file ", outputfile ," has been created with", total_num_lines, "lines"
