# -*- coding:utf-8 -*-
#
# author:qingehua
# Python3
# 功能：清理大V微博内容及评论中的@用户名
# 输入文件：第36行
# 输出文件：第40行
# 输入文件格式如下：
# 1 ID | 2 微博名 | 3 微博地址 | 4 性别 | 5 省份 | 6 微博总数 | 7 关注数 | 8 粉丝数 | 9 微博发送时间 | 10 赞 | 11 转发 | 12 评论 | 13 微博内容 - 评论

import time            
import re            
import os    
import sys  
import codecs
import xlrd
import xlsxwriter

def clearAT(content):
	clearedContent = []
	flag = 0
	if not isinstance(content, str):
		return content
	for i in content:
		if i == '@':
			flag = 1
			continue
		elif i != ':' and flag == 1:
			continue
		elif i == ':' and flag == 1:
			flag = 0
			continue
		clearedContent.append(i)

	return ''.join(clearedContent)

input_name = 'VWeibo.xlsx' 
readbook = xlrd.open_workbook(input_name)
sheet1 = readbook.sheet_by_index(0)

output_name = 'VWeibo_clearAT.xlsx'
writebook = xlsxwriter.Workbook(output_name)
sheet_info = writebook.add_worksheet()

# 1 ID | 2 微博名 | 3 微博地址 | 4 性别 | 5 省份 | 6 微博总数 | 7 关注数 | 8 粉丝数 | 9 微博发送时间 | 10 赞 | 11 转发 | 12 评论 | 13 微博内容 - 评论

for temp_row in range(sheet1.nrows):
	if temp_row == 0:
		continue
	sheet_info.write(temp_row,0,sheet1.row(temp_row)[0].value)				
	sheet_info.write(temp_row,1,sheet1.row(temp_row)[1].value)															
	sheet_info.write(temp_row,2,sheet1.row(temp_row)[2].value)											
	sheet_info.write(temp_row,3,sheet1.row(temp_row)[3].value)											
	sheet_info.write(temp_row,4,sheet1.row(temp_row)[4].value)											
	sheet_info.write(temp_row,5,sheet1.row(temp_row)[5].value)											
	sheet_info.write(temp_row,6,sheet1.row(temp_row)[6].value)											
	sheet_info.write(temp_row,7,sheet1.row(temp_row)[7].value)											
	sheet_info.write(temp_row,8,sheet1.row(temp_row)[8].value)											
	sheet_info.write(temp_row,9,sheet1.row(temp_row)[9].value)											
	sheet_info.write(temp_row,10,sheet1.row(temp_row)[10].value)
	sheet_info.write(temp_row,11,sheet1.row(temp_row)[11].value)

	for i in range(len(sheet1.row_values(temp_row))-12):
		sheet_info.write(temp_row,i+12,clearAT(sheet1.row(temp_row)[i+12].value))
		#print(clearAT(sheet1.row(temp_row)[i+12].value))
													
	print(temp_row)

writebook.close()