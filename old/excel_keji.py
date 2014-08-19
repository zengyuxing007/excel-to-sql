# -*- coding=UTF-8 -*-
import xlrd
data = xlrd.open_workbook('./2.xls') 
sheet = data.sheets()[1]  
nrows = sheet.nrows  
ncols = sheet.ncols 
i = 0
x = 0
cells = {
	'no' : ' ',
	'college' : ' ',
	'works_name' : ' ',
	'type' : ' ',
	'teamleader' : ' ',
	'teamleader_class' : ' ',
	'teacher' : ' ',
        'wang_money': ' ',
        'end_money' : ' ',
}
while i<nrows:
        while x<ncols:
                if x == 0:
                        cells['no'] = sheet.cell(i,x).value
                elif x == 1:
                        cells['works_name'] = sheet.cell(i,x).value
                elif x == 2:
                        cells['teamleader'] = sheet.cell(i,x).value
                elif x == 3:
                        cells['teamleader_class'] = sheet.cell(i,x).value
                elif x == 4:
                        cells['college'] = sheet.cell(i,x).value
                elif x == 5:
                        cells['teacher'] = sheet.cell(i,x).value
                elif x == 6:
                        cells['type'] = sheet.cell(i,x).value
                elif x == 7:
                        cells['wang_money'] = unicode(sheet.cell(i,x).value)
                elif x == 8:
                        cells['end_money'] = unicode(sheet.cell(i,x).value)
                x += 1
        sk = u'('+ u"'" + cells['no']+ u"'" + u',' + ' ' + u"'" + cells['works_name'] + u"'" + u',' + ' ' +  u"'" + cells['teamleader'] + u"'" + u','+ ' ' + u"'" + cells['teamleader_class'] + u"'" + u','+ ' ' + u"'" + cells['college'] + u"'" + u','+ ' ' + u"'" + cells['teacher'] + u"'" + u','+ u"'" + cells['type'] + u"'" + u','+ ' ' + u"'" + cells['wang_money'] + u"'" + u','+ ' ' + u"'" + cells['end_money'] + u"'" + u')' + u','#+ ' ' + u"'" + cells['teacher'] + u"'" + u','+ ' ' + u"'" + str(cells['similarity']) + u'%' + u"'" + u','+ ' ' + str(cells['written_grade']) + u',' + ' ' + str(cells['reply_grade']) + u',' + ' ' + str(cells['end_grade']) + u','+ ' ' + u"'" + cells['grade_school'] + u"'" + u','+ ' ' + u"''" + u','+ ' ' + u"'" + u'校' + u"'" + u')' + u','
        sk2 = sk.encode('utf-8')
        f = open('./q', 'a')
        f.write(sk2)
        br = '\n'
        f.write(br)
        f.close()
        i +=1
        x = 0
print 'ok'