
# open General Ledger from Excel, reformat into a database. 
import pandas as pd 
import xlsxwriter as xlwrit
from pathlib import Path

col_name = ["Account Description","Type","Date","Num","Name","Memo","Split","Amount","Balance"]

#open excel, add to dataframe, drop blank columns/rows
def format_file(xcel):

	#Create merged column from Col 1 & Col 2, and delete cols when finished
	col1_null = xcel[1].isnull()
	col2_null = xcel[2].isnull()
	col3_null = xcel[3].isnull()
	newCol = []

	for x in range(col1_null.size):

		if not col1_null[x]:
			newCol.append(str(xcel.iloc[x, 1]))
		elif not col2_null[x]:
			newCol.append(str(xcel.iloc[x, 2]))
		elif not col3_null[x]:
			newCol.append(str(xcel.iloc[x, 3]))
		elif x > 0:
			newCol.append(newCol[x-1])
		else:
			newCol.append("")



	xcel['Account Description'] = newCol

	#clean up empty columns/rows
	xcel = xcel.drop([0,1,2,3,4], axis = 1)
	xcel = xcel.dropna(axis = 0, how='all')
	xcel = xcel.dropna(axis = 1, how='all')

	#remove top row (headers) & total rows
	xcel = xcel.drop(0)
	xcel = xcel.dropna(axis = 0, thresh = 6)

	#Reindex Columns
	cols = xcel.columns.tolist()
	cols = cols[-1:] + cols[:-1]
	xcel = xcel[cols]

	xcel.columns = col_name

	#Write to excel file
	writer = pd.ExcelWriter("modified GL.xlsx", engine='xlsxwriter')
	xcel.to_excel(writer,sheet_name='Sheet1')

	wb = writer.book
	ws = writer.sheets['Sheet1']
	ws.set_column('A:J', 15)

	wb.close()

def open_file(filename):
	try:
		 x = pd.read_excel(filename, index_col = None, header = None)
		 return x
	except:
		input("Error with filename.\n Press any key to exit")
		quit()

def main():
	format_file(open_file(input("Enter Filename \n")))
	input("Hit any key to continue...")

#Run program
if __name__ == "__main__":
	main()
