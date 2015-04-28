#-*- coding:utf-8 -*-
import xlrd
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

####打开excel表格####
book = xlrd.open_workbook("42examination.xls")
sheets_book = book.sheets()  ##获取excel额工作簿数目

####输出本次工作簿的考务注意事项.
table = book.sheets()[0]
import_notice = table.cell_value(25, 1)

####
languageset = [15, 24, 26, 27, 28, 29, 35, 36, 37, 38, 39, 41, 42, 61, 63, 64, 65]
examniation_type = [u'一级MS OFFICE', u'二级C', u'二级VISUAL BASIC', u'二级VISUAL FOXPRO', 
	u'二级JAVA', u'二级ACCESS', u'三级网络技术', u'三级数据库技术', u'三级软件测试技术', u'三级信息安全技术', 
	u'三级嵌入式系统开发技术',u'四级网络工程师', u'四级数据库工程师', u'二级C++', 
	u'二级MySQL数据库程序设计', u'二级Web程序设计', u'二级MS Office高级应用']

#####
def find_cell_total_people():
	for colnun in range(table.ncols):
		comm_cell = table.row(5)[colnun]
		##设置单元格格式去通配内容
		if comm_cell.value == '考生总人数':
			print '每个科目考生总人数在第%d列' %(colnun)
	return colnun
total_candidates = find_cell_total_people() ##测试函数

###获取每个类别考试科目的考生人数
def get_total(base_col):
	kemu_total_per = []
	while base_col < (table.nrows - 2):
		per = table.col(total_candidates)[base_col].value
		kemu_total_per.append(per)
		base_col += 1
	return kemu_total_per
total_examinations = get_total(7)

def info_show_from_excel():
	for i, y, total in zip(range(len(languageset)), range(len(examniation_type)), range(len(total_examinations))):
		# print "%s 的考试语言代码是 %s", %(examniation_type[y].encode('utf-8'), languageset[i])
		return "%s   考场语言代码:   %d , 该科目考生人数是: ## %s ##" % (examniation_type[y].encode('utf-8') , languageset[i], total_examinations[total])