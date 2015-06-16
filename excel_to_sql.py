# coding=utf-8
import sys
import xlrd
import re

reload(sys)
sys.setdefaultencoding('utf-8')

filename = raw_input('Input the original excel filename: ')

data = xlrd.open_workbook(filename)

with open('insert.sql', 'w') as out_file:
	for sheet in data.sheets():
		table_name = sheet.name
		cols = map(lambda item: ('[' + item[:-4] + ']') if item.endswith('.NaN') else ('[' + item + ']'), sheet.row_values(0))
		nan_flags = map(lambda item: True if item.endswith('.NaN') else False, sheet.row_values(0))

		out_file.write('-- ' + table_name + '\r\n')
		for i in range(1,sheet.nrows):
			row_values = sheet.row_values(i)
			for j in range(0,len(row_values)):
				if nan_flags[j] or not re.match(r'^[\d\.]+$', str(row_values[j])):
					row_values[j] = '\'' + str(row_values[j]) + '\''
			insert_sql = 'INSERT INTO [%s] (%s) VALUES(%s);' % (table_name, ', '.join(cols), ', '.join(row_values))
			out_file.write(insert_sql + '\r\n')
print 'SQL file was gernarated successfully! See insert.sql.'
