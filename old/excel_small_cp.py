# -*- coding=UTF-8 -*-
import xlrd
data = xlrd.open_workbook('./2.xls') 
sheet = data.sheets()[1]  
# nrows = sheet.nrows  
# ncols = sheet.ncols 
# i = 0
# x = 0

# cells = {
# 	'ranking' : ' ',
# 	'no' : ' ',
# 	'college' : ' ',
# 	'works_name' : ' ',
# 	'teams_name' : ' ',
# 	'type' : ' ',
# 	'teamleader' : ' ',
# 	'teamleader_class' : ' ',
# 	'oth_name' : ' ',
# 	'teacher' : ' ',
# 	'similarity' : ' ',
# 	'written_grade' : ' ',
# 	'reply_grade' : ' ',
# 	'end_grade' : ' ',
# 	'grade_school' : ' '
# }
# while i<nrows:
#   while x<ncols:
#     if x == 0:
#             cells['ranking'] = unicode(round(sheet.cell(i,x).value,0))
#     elif x == 1:
#             cells['no'] = sheet.cell(i,x).value
#     elif x == 2:
#             cells['college'] = sheet.cell(i,x).value
#     elif x == 3:
#             cells['works_name'] = sheet.cell(i,x).value
#     elif x == 4:
#             cells['teams_name'] = sheet.cell(i,x).value
#     elif x == 5:
#             cells['type'] = sheet.cell(i,x).value
#     elif x == 6:
#             cells['teamleader'] = sheet.cell(i,x).value
#     elif x == 7:
#             cells['teamleader_class'] = sheet.cell(i,x).value
#     elif x == 8:
#             cells['oth_name'] = sheet.cell(i,x).value
#     elif x == 9:
#             cells['teacher'] = sheet.cell(i,x).value
#     elif x == 10:
#             cells['similarity'] = unicode(round(sheet.cell(i,x).value*100,4))
#     elif x == 11:
#             cells['written_grade'] = unicode(round(sheet.cell(i,x).value,2))
#     elif x == 12:
#             cells['reply_grade'] = unicode(round(sheet.cell(i,x).value,2))
#     elif x == 13:
#             cells['end_grade'] = unicode(round(sheet.cell(i,x).value,2))
#     elif x == 14:
#             cells['grade_school'] = sheet.cell(i,x).value
#     x += 1
#   sk = u'(' + cells['ranking'] + u',' + ' ' + u"'" + cells['no'] + u"'" + u',' + ' ' +  u"'" + cells['college'] + u"'" + u','+ ' ' + u"'" + cells['works_name'] + u"'" + u','+ ' ' + u"'" + cells['teams_name'] + u"'" + u','+ ' ' + u"'" + cells['type'] + u"'" + u',' + ' ' + u"''" + u',' + ' ' + u"'" + cells['teamleader'] + u"'" + u','+ ' ' + u"'" + cells['teamleader_class'] + u"'" + u','+ ' ' + u"'" + cells['oth_name'] + u"'" + u','+ ' ' + u"'" + cells['teacher'] + u"'" + u','+ ' ' + u"'" + str(cells['similarity']) + u'%' + u"'" + u','+ ' ' + str(cells['written_grade']) + u',' + ' ' + str(cells['reply_grade']) + u',' + ' ' + str(cells['end_grade']) + u','+ ' ' + u"'" + cells['grade_school'] + u"'" + u','+ ' ' + u"''" + u','+ ' ' + u"'" + u'校' + u"'" + u')' + u','
#   sk2 = sk.encode('utf-8')
#   f = open('./q', 'a')
#   f.write(sk2)
#   br = '\n'
#   f.write(br)
#   f.close()
#   i +=1
#   x = 0
# print 'ok'
