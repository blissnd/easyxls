from easyxls import *

### Pivot Table example
data_struct = get_spreadsheet(spreadsheet_path='homers_shopping_list.xlsx', max_row=13, max_column='g', header_row_start=8, header_col_start='c', format='pivot')

### Column headings based example
#data_struct = get_spreadsheet(spreadsheet_path='janendra.xlsx', max_row=13, max_column='g', header_row_start=8, header_col_start='d', format='column')

### Row headings based example
#data_struct = get_spreadsheet(spreadsheet_path='janendra.xlsx', max_row=13, max_column='g', header_row_start=9, header_col_start='c', format='row')

#########

if type(data_struct) == type([]):
	for elem in data_struct:
		print elem
		print
	# End For
# End If

if type(data_struct) == type({}):
	for key,val in data_struct.items():
		print key + ' => ' + str(val)
		print
	# End For
# End If
