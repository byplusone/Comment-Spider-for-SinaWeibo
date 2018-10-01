# -*- coding:utf-8 -*-
#
# author: qingehua
# 功能：将爬好的大V评论数据、评论用户微博进行词频统计，并根据词汇库进行各值的计算
# 输入格式：
# 		大V评论（**注意这里输入没有首行题头**）：
# 			1 ID | 2 微博名 | 3 微博地址 | 4 性别 | 5 省份 | 6 微博总数 | 7 关注数 | 8 粉丝数 | 9 微博发送时间 | 10 赞 | 11 转发 | 12 评论 | 13 微博内容 - 评论
#		评论用户（**直接按照爬虫输出格式即可**）：
#			1 ID | 2 微博名 | 3 微博地址 | 4 性别 | 5 省份 | 6 生日 | 7 认证信息 | 8 微博总数 | 9 关注数 | 10 粉丝数 | 9 微博发送时间 | 10 微博内容 | 11 赞 | 12 转发 | 13 评论 - 评论
#输出格式：
#		第一页（不加权）：词频 | value1 | value2 | value3 | …
#		第二页（加权）：词频 | value1 | value2 | value3 | …
#
# 计算大V和计算评论用户的代码是分开的，请在注释掉相应部分后再计算
#
import time            
import re            
import os    
import sys  
import codecs  
import shutil
import xlrd
import xlsxwriter

def loadVec():
	wordVec = open('words_vec.txt', 'r')
	line = wordVec.readline()
	Vec = []
	while(line):
		keywords = str(line).split(';')
		keywordID = keywords[0]
		keyword = keywords[1]
		value1 = float(keywords[2])
		value2 = float(keywords[3])
		value3 = float(keywords[4])
		Vec.append([keywordID, keyword, value1, value2, value3])
		line = wordVec.readline()
	return Vec

def ugMining(sourceName, outputName):
	wordVec = open('words_vec.txt', 'r')

	writebook = xlsxwriter.Workbook(outputName)
	ugResult = writebook.add_worksheet()
	weightedResult = writebook.add_worksheet()
	line = wordVec.readline()

	readFile = xlrd.open_workbook(sourceName)
	#readFile = xlrd.open_workbook('test.xlsx')
	readSheet = readFile.sheet_by_index(0)
	count = 1

	while(line):
		print(count)
		keywords = str(line).split(';')
		keywordID = keywords[0]
		keyword = keywords[1]
		value1 = float(keywords[2])
		value2 = float(keywords[3])
		value3 = float(keywords[4])

		addedUsers = []
		addedUsersCount = []
		addedUsersWeights = []
		for i in range(readSheet.nrows):
			if i == 0:
				continue
			temp_user = str(readSheet.row(i)[1])
			temp_user = temp_user[6:16]

			if temp_user not in addedUsers:
				addedUsers.append(temp_user)
				addedUsersCount.append(0)

				addedUsersWeight = str(readSheet.row(i)[8])
				addedUsersWeight = re.findall(r"\d+", addedUsersWeight, re.M)
				addedUsersWeight = int(addedUsersWeight[0])
				addedUsersWeights.append(addedUsersWeight)

			temp_user_loc = addedUsers.index(temp_user)

			temp_result = re.findall(keyword, str(readSheet.row(i)[12]))
			temp_result = len(temp_result)
			addedUsersCount[temp_user_loc] += temp_result

		
		ugResult.write(count,0,keywordID)
		ugResult.write(count,1,keyword)

		weightedResult.write(count,0,keywordID)
		weightedResult.write(count,1,keyword)

		for i in range(len(addedUsers)):
			ugResult.write(0, 4*i+2, 'Frequency Count: ' + addedUsers[i])
			ugResult.write(0, 4*i+3, '正负度: ' + addedUsers[i])
			ugResult.write(0, 4*i+4, '唤起度: ' + addedUsers[i])
			ugResult.write(0, 4*i+5, '支配度: ' + addedUsers[i])

			ugResult.write(count, 4*i+2, str(addedUsersCount[i]))
			ugResult.write(count, 4*i+3, str(addedUsersCount[i]*value1))
			ugResult.write(count, 4*i+4, str(addedUsersCount[i]*value2))
			ugResult.write(count, 4*i+5, str(addedUsersCount[i]*value3))

			####################### Weighted Version ######################

			weightedResult.write(0, 4*i+2, 'Weighted Frequency Count: ' + addedUsers[i])
			weightedResult.write(0, 4*i+3, '加权正负度: ' + addedUsers[i])
			weightedResult.write(0, 4*i+4, '加权唤起度: ' + addedUsers[i])
			weightedResult.write(0, 4*i+5, '加权支配度: ' + addedUsers[i])

			weightedResult.write(count, 4*i+2, str(addedUsersCount[i]*addedUsersWeights[i]))
			weightedResult.write(count, 4*i+3, str(addedUsersCount[i]*value1*addedUsersWeights[i]))
			weightedResult.write(count, 4*i+4, str(addedUsersCount[i]*value2*addedUsersWeights[i]))
			weightedResult.write(count, 4*i+5, str(addedUsersCount[i]*value3*addedUsersWeights[i]))

		line = wordVec.readline()
		count += 1

	writebook.close()


def VMining(sourceName, outputName):
	writebook = xlsxwriter.Workbook(outputName)
	VResult = writebook.add_worksheet()
	weightedVResult = writebook.add_worksheet()
	#line = wordVec.readline()

	readFile = xlrd.open_workbook(sourceName)
	readSheet = readFile.sheet_by_index(0)
	wordVec = loadVec()

	for i in range(readSheet.nrows):
		print(i)
		# 注意这里第一行一定不为空或者题头
		temp_user = str(readSheet.row(i)[0])
		temp_user = temp_user[6:11]

		addedUsersWeight = str(readSheet.row(i)[11])
		addedUsersWeight = re.findall(r"\d+", addedUsersWeight, re.M)
		addedUsersWeight = int(addedUsersWeight[0])

		VResult.write(0, 4*i+2, 'Frequency Count: ' + temp_user)
		VResult.write(0, 4*i+3, '正负度: ' + temp_user)
		VResult.write(0, 4*i+4, '唤起度: ' + temp_user)
		VResult.write(0, 4*i+5, '支配度: ' + temp_user)

		weightedVResult.write(0, 4*i+2, 'Weighted Frequency Count: ' + temp_user)
		weightedVResult.write(0, 4*i+3, '加权正负度: ' + temp_user)
		weightedVResult.write(0, 4*i+4, '加权唤起度: ' + temp_user)
		weightedVResult.write(0, 4*i+5, '加权支配度: ' + temp_user)

		for keyNum, keyVec in enumerate(wordVec):
			addedUsersCount = 0
			for j in range(len(readSheet.row_values(i))-12):
				addedUsersCount += len(re.findall(keyVec[1], str(readSheet.row(i)[j+12])))

			VResult.write(keyNum+1, 4*i+2, str(addedUsersCount))
			VResult.write(keyNum+1, 4*i+3, str(addedUsersCount*keyVec[2]))
			VResult.write(keyNum+1, 4*i+4, str(addedUsersCount*keyVec[3]))
			VResult.write(keyNum+1, 4*i+5, str(addedUsersCount*keyVec[4]))

			weightedVResult.write(keyNum+1, 4*i+2, str(addedUsersCount*addedUsersWeight))
			weightedVResult.write(keyNum+1, 4*i+3, str(addedUsersCount*keyVec[2]*addedUsersWeight))
			weightedVResult.write(keyNum+1, 4*i+4, str(addedUsersCount*keyVec[3]*addedUsersWeight))
			weightedVResult.write(keyNum+1, 4*i+5, str(addedUsersCount*keyVec[4]*addedUsersWeight))

			print('--' + keyVec[0])
	

	for keyNum, keyVec in enumerate(wordVec):
		VResult.write(keyNum+1, 0, keyVec[0])
		VResult.write(keyNum+1, 1, keyVec[1])
		weightedVResult.write(keyNum+1, 0, keyVec[0])
		weightedVResult.write(keyNum+1, 1, keyVec[1])

	writebook.close()

def joinCmmts(sourceName):
	writebook = xlsxwriter.Workbook('qgh_temp.xlsx')
	joinedCmmtSheet = writebook.add_worksheet()

	readFile = xlrd.open_workbook(sourceName)
	readSheet = readFile.sheet_by_index(0)
	
	for i in range(readSheet.nrows):
		for j in range(12):
			joinedCmmtSheet.write(i, j, str(readSheet.row(i)[j].value))
		text = ''
		for j in range(len(readSheet.row_values(i))-12):
			text += str(readSheet.row(i)[j+12].value)
		joinedCmmtSheet.write(i, 12, text)
	print('Comments joined!')
	writebook.close()




def advancedVMining(sourceName, outputName):
	writebook = xlsxwriter.Workbook(outputName)
	VResult = writebook.add_worksheet()
	weightedVResult = writebook.add_worksheet()
	#line = wordVec.readline()

	# 在存在qgh_temp.xlsx的情况下可以注释掉这一句，节约初始化时间
	joinCmmts(sourceName)

	readFile = xlrd.open_workbook('qgh_temp.xlsx')
	readSheet = readFile.sheet_by_index(0)
	wordVec = loadVec()

	for i in range(readSheet.nrows):
		print(i)
		# 注意这里第一行一定不为空或者题头
		temp_user = str(readSheet.row(i)[0])
		temp_user = temp_user[6:11]

		addedUsersWeight = str(readSheet.row(i)[11])
		#print(addedUsersWeight)
		addedUsersWeight = re.findall(r"\d+", addedUsersWeight, re.M)
		addedUsersWeight = int(addedUsersWeight[0])

		VResult.write(0, 4*i+2, 'Frequency Count: ' + temp_user)
		VResult.write(0, 4*i+3, '正负度: ' + temp_user)
		VResult.write(0, 4*i+4, '唤起度: ' + temp_user)
		VResult.write(0, 4*i+5, '支配度: ' + temp_user)

		weightedVResult.write(0, 4*i+2, 'Weighted Frequency Count: ' + temp_user)
		weightedVResult.write(0, 4*i+3, '加权正负度: ' + temp_user)
		weightedVResult.write(0, 4*i+4, '加权唤起度: ' + temp_user)
		weightedVResult.write(0, 4*i+5, '加权支配度: ' + temp_user)


		for keyNum, keyVec in enumerate(wordVec):
			addedUsersCount = len(re.findall(keyVec[1], str(readSheet.row(i)[12])))

			VResult.write(keyNum+1, 4*i+2, str(addedUsersCount))
			VResult.write(keyNum+1, 4*i+3, str(addedUsersCount*keyVec[2]))
			VResult.write(keyNum+1, 4*i+4, str(addedUsersCount*keyVec[3]))
			VResult.write(keyNum+1, 4*i+5, str(addedUsersCount*keyVec[4]))

			weightedVResult.write(keyNum+1, 4*i+2, str(addedUsersCount*addedUsersWeight))
			weightedVResult.write(keyNum+1, 4*i+3, str(addedUsersCount*keyVec[2]*addedUsersWeight))
			weightedVResult.write(keyNum+1, 4*i+4, str(addedUsersCount*keyVec[3]*addedUsersWeight))
			weightedVResult.write(keyNum+1, 4*i+5, str(addedUsersCount*keyVec[4]*addedUsersWeight))

			#print('--' + keyVec[0])
	

	for keyNum, keyVec in enumerate(wordVec):
		VResult.write(keyNum+1, 0, keyVec[0])
		VResult.write(keyNum+1, 1, keyVec[1])
		weightedVResult.write(keyNum+1, 0, keyVec[0])
		weightedVResult.write(keyNum+1, 1, keyVec[1])


	writebook.close()


if __name__ == '__main__':
	#################################################
	#
	#
	#				计算评论用户微博
	#
	sourceFileName = 'L0001_Cmmt300_60-79.xlsx'
	ugFileName = 'L0001_Cmmt300_60-79_mining.xlsx'
	ugMining(sourceFileName, ugFileName)
	#################################################
	#
	#
	#				计算大V评论
	#
	#sourceFileName = 'VWeibo.xlsx'
	#VFileName = 'VWeibo_mining.xlsx'
	#VMining(sourceFileName, VFileName) # 旧版函数，效率较慢
	#advancedVMining(sourceFileName, VFileName)