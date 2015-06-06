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
		cols = map(lambda item: '`' + item + '`', sheet.row_values(0))
		out_file.write('-- ' + table_name + '\r\n')
		for i in range(1,sheet.nrows):
			items = map(lambda item: str(item.value) if re.match(r'^[\d\.]+$', str(item.value)) else '\'' + str(item.value) + '\'', sheet.row(i))
			insert_sql = 'INSERT INTO %s (%s) VALUES(%s);' % (table_name, ', '.join(cols), ', '.join(items))
			out_file.write(insert_sql + '\r\n')
print 'SQL file was gernarated successfully! See insert.sql.'