#!/usr/bin/python 
#-*- coding: UTF-8 -*-

import xlrd
import json
import sys

# load config.json error processing
try:
	config = json.loads(open('./config.json','r').read())
except ValueError as inst:
	print('\n    config.json format error.')
	print('    ' + (inst.args[0]) + '.')
	sys.exit()
except IOError as inst:
	print('\n    config.json not found.')
	print('    ' + (str(inst))[10:] + '.')
	sys.exit()

# check config.json data type
try:
	int(config['excel_sheet_number'])    # number
except ValueError as inst:
	print('\n    \'excel_sheet_number\' only allow Number ,pleace check the data type in config.json.')
	sys.exit()

try:
	for keys_data_type in config['cols_info']:   # number
		int(keys_data_type)
except ValueError as inst:
	print('\n    \'cols_info\' keys only allow Number ,pleace check the data type in config.json.')
	sys.exit()

try:
	len(config['start_row']) == 0 or int(config['start_row'])   # number none
except ValueError as inst:
	print('\n    \'start_row\' only allow Number or "" ,pleace check the data type in config.json.')
	sys.exit()

try:
	len(config['end_row']) == 0 or int(config['end_row'])   # number none
except ValueError as inst:
	print('\n    \'end_row\' only allow Number or "" ,pleace check the data type in config.json.')
	sys.exit()

try:
	for barring_row in config['barring_row']:  # number none
		int(barring_row)
except ValueError as inst:
	print('\n    \'barring_row\' only allow Number or "" ,pleace check the data type in config.json.')
	sys.exit()

# load xls error processing
try:
	xls = xlrd.open_workbook(config['src'])
except xlrd.biffh.XLRDError as inst:
	print('\n    ' + config['src'] + ' load error, please check the xls whether can be open.')
	print('    ' + (inst.args[0]) + '.')
	sys.exit()
except IOError as inst:
	print('\n    ' + (str(inst))[10:] + '.')
	print('    please check \"src\" in config.json')
	sys.exit()

#db sheet index
if len(config['db_sheet_name']) == 0:
	print('\n    you should set a \'db_sheet_name\' in config.json.')
	sys.exit()

# initialize arguments
# set all index 


num = int(config['excel_sheet_number']) -1 
sheet = xls.sheets()[num]
nrows = sheet.nrows
ncols = sheet.ncols

# if don't setup ,default start from row 0.
if len(config['start_row']) == 0:
	config['start_row'] = 0
elif int(config['start_row']) <= 0:
	config['start_row'] = 0
elif int(config['start_row']) >= nrows:
	config['start_row'] = 0
else:
	config['start_row'] = int(config['start_row']) - 1

# if don't setup ,default end up with last row. 
if len(config['end_row']) == 0:
	config['end_row'] = nrows - 1
elif int(config['end_row']) > nrows:
	config['end_row'] = nrows - 1
elif int(config['end_row']) <= 0:
	config['end_row'] = nrows - 1
elif int(config['end_row']) <= int(config['start_row']):
	print('\n    \'end_row\' shoule more than \'start_row\'')
	sys.exit()
else:
	config['end_row'] = int(config['end_row']) - 1

# barring_row
n = len(config['barring_row']) - 1
if n < 0:
	config['barring_row'] = None
else:
	while n>=0:
		if int(config['barring_row'][n]) <= 0:
			config['barring_row'][n] = 0
		elif int(config['barring_row'][n]) > nrows:
			config['barring_row'][n] = nrows - 1
		else:
			config['barring_row'][n] = int(config['barring_row'][n]) - 1
		n = n - 1

# start ergodic
i = int(config['start_row'])
e = int(config['end_row'])

sql = "INSERT INTO `" + config['db_sheet_name'] + "` ("

for fields in config['cols_info']:
	sql = sql + "`" + config['cols_info'][fields][0] + "`,"
sql = sql[:int(len(sql))-1] + ') VALUES \n'

detail_data = ''

if config['barring_row'] == None:
	while i<=e:
		detail_data = detail_data + u'('
		for cols in config['cols_info']:
			syn = int(cols)-1
			if config['cols_info'][cols][1] == 'num':
				detail_data = detail_data + unicode(sheet.cell(i,syn).value) + u','
			elif config['cols_info'][cols][1] == 'str':
				detail_data = detail_data + u"'" + unicode(sheet.cell(i,syn).value) + u"'" + u','
		detail_data = detail_data[:int(len(detail_data))-1]
		detail_data = detail_data + u'),\n'
		i+=1
else:
	while i<=e:
		if i in config['barring_row']:
			pass
		else:
			detail_data = detail_data + u'('
			for cols in config['cols_info']:
				syn = int(cols)-1
				if config['cols_info'][cols][1] == 'num':
					detail_data = detail_data + unicode(sheet.cell(i,syn).value) + u','
				elif config['cols_info'][cols][1] == 'str':
					detail_data = detail_data + u"'" + unicode(sheet.cell(i,syn).value) + u"'" + u','
			detail_data = detail_data[:int(len(detail_data))-1]
			detail_data = detail_data + u'),\n'
		i+=1


detail_data = detail_data[:int(len(detail_data))-2] + u';'

res = (sql + detail_data).encode('utf-8')

f = open(config['output'], 'w')
f.write(res)
f.close()
