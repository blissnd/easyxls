from easyxls import *

### Pivot Table example
#data_struct = get_spreadsheet(spreadsheet_path='homers_shopping_list.xlsx', max_row=40, max_column='ad', header_row_start=35, header_col_start='z', format='pivot')

### Column headings based example
#data_struct = get_spreadsheet(spreadsheet_path='homers_shopping_list.xlsx', max_row=40, max_column='ad', header_row_start=35, header_col_start='aa', format='column')

### Row headings based example
data_struct = get_spreadsheet(spreadsheet_path='homers_shopping_list.xlsx', max_row=40, max_column='ad', header_row_start=36, header_col_start='z', format='row')

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
