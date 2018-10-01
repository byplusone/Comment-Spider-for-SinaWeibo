# -*- coding:utf-8 -*-
#
# author: qingehua
# 功能：把词汇库转为txt格式
import time            
import re            
import os    
import sys  
import codecs  
import shutil
import xlrd
import xlsxwriter

outputFile = open('words_vec.txt', 'w')
readFile = xlrd.open_workbook('words_vec.xls')
readSheet = readFile.sheet_by_index(0)

for i in range(readSheet.nrows):
	data1 = str(readSheet.row(i)[0]).split('\'')
	data1 = data1[1]
	data2 = str(readSheet.row(i)[1]).split('\'')
	data2 = data2[1]
	data3 = str(readSheet.row(i)[2]).split(':')
	data3 = data3[1]
	data4 = str(readSheet.row(i)[3]).split(':')
	data4 = data4[1]
	data5 = str(readSheet.row(i)[4]).split(':')
	data5 = data5[1]
	outputFile.write(data1 + ';' + data2+  ';' + data3 + ';' + data4 + ';' + data5 + '\n')