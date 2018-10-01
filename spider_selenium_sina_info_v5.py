# coding=utf-8
import time            
import re            
import os    
import sys  
import codecs  
import shutil
import urllib 
from selenium import webdriver        
from selenium.webdriver.common.keys import Keys        
import selenium.webdriver.support.ui as ui        
from selenium.webdriver.common.action_chains import ActionChains
import xlrd
import xlsxwriter

def clearAT(content):
	clearedContent = []
	flag = 0
	for i in content:
		if i == u'@':
			flag = 1
			continue
		elif i != u' ' and flag == 1:
			continue
		elif i == u' ' and flag == 1:
			flag = 0
			continue
		clearedContent.append(i)
	return ''.join(clearedContent)


#先调用无界面浏览器PhantomJS或Firefox    
driver = webdriver.Firefox()
wait = ui.WebDriverWait(driver,10)

user_name = u'L0001'

#全局变量 文件操作读写信息
output_name = user_name + '_Cmmt300.xlsx' 
writebook = xlsxwriter.Workbook(output_name)
sheet_info = writebook.add_worksheet()
comment_user_info = writebook.add_worksheet()
temp_row = 1 # 定义写操作时全局的当前操作的行数

def LoginWeibo(username, password):
    try:
        print u'准备登陆Weibo.cn网站...' 
        driver.get("https://passport.weibo.cn/signin/login?entry=mweibo&r=http%3A%2F%2Fweibo.cn%2F&backTitle=%CE%A2%B2%A9&vt=")
        time.sleep(40)
        print u'登陆成功...'
        
    except Exception,e:      
        print "Error: ",e
    finally:    
        print u'End LoginWeibo!\n\n'

def VisitLeafPage(web_url, innner_id):
	try:
		global temp_row

		print u'叶节点用户'+str(innner_id)
		driver.get(web_url)

		#昵称
		str_name = driver.find_element_by_xpath("//div[@class='ut']")
		str_t = str_name.text.split(" ")
		num_name = str_t[0]      #空格分隔 获取第一个值 "Eastmount 详细资料 设置 新手区"
		#print u'昵称: ' + num_name 

		str_wb = driver.find_element_by_xpath("//div[@class='tip2']")  
		pattern = r"\d+\.?\d*"   #正则提取"微博[0]" 但r"(\[.*?\])"总含[] 
		guid = re.findall(pattern, str_wb.text, re.S|re.M)
		#print str_wb.text        #微博[294] 关注[351] 粉丝[294] 分组[1] @他的
		for value in guid:
			num_wb = int(value)
			break
		#print u'微博数: ' + str(num_wb)

		#关注数
		str_gz = driver.find_element_by_xpath("//div[@class='tip2']/a[1]")
		guid = re.findall(pattern, str_gz.text, re.M)
		num_gz = int(guid[0])
		#print u'关注数: ' + str(num_gz)

		#粉丝数
		str_fs = driver.find_element_by_xpath("//div[@class='tip2']/a[2]")
		guid = re.findall(pattern, str_fs.text, re.M)
		num_fs = int(guid[0])
		#print u'粉丝数: ' + str(num_fs)

		#info页面
		#获取到新浪微博官方为用户设计的id
		str_fs = driver.find_element_by_xpath("//div[@class='tip2']/a[2]")
		temp_info_id = str_fs.get_attribute("href")
		temp_info_id = temp_info_id.split('/')
		info_id = temp_info_id[3]
		info_web = "https://weibo.cn/" + info_id + "/info"
		#print info_web
		driver.get(info_web)
		basic_info = driver.find_element_by_xpath("//div[6]")
		basic_info = basic_info.text.split('\n')
		#print basic_info
		
		sex = 'n/a'
		loc = 'n/a'
		birth = 'n/a'
		uid = 'n/a'
		quali_id = 'n/a'
		print "SAFE HERE"
		for item in basic_info:
			if u"昵称" in item:
				uid = item.split(':')
				uid = uid[1]
			if u"性别" in item:
				sex = item.split(':')
				sex = sex[1]
			if u"地区" in item:
				loc = item.split(':')
				loc = loc[1]
				loc = loc.split(' ')[0]
			if u"生日" in item:
				birth = item.split(':')
				birth = birth[1]
			if u"认证信息" in item:
				quali_id = item.split(u'：')
				quali_id = quali_id[1]
				break
	    

		#进入原创微博页面
		original_weibo = "https://weibo.cn/u/" + str(info_id) + "?filter=1&page=1"
		print "IM SAFE HERE: " + original_weibo
		#https://weibo.cn/5623115983?filter=1&page=3
		driver.get(original_weibo)
		info = driver.find_elements_by_xpath("//div[@class='c']")
		
		page_num = 1 # 记录当前用户微博的第n页
		weibo_num = 1 # 记录当前用户的第n条微博

		while weibo_num <= 300:
			#time.sleep(2)
			print "VISITING USER PAGE #" + str(page_num)
			if len(info) < 3:
				break

			for i,value in enumerate(info):

				if i > len(info)-3:
					break;				
				content = value.text

				str1 = content.split(u"赞[")[1]
				str1 = str1.split("]")[0]
				likes = str1

				str2 = content.split(u"转发[")[1]
				str2 = str2.split("]")[0]
				trans = str2

				str3 = content.split(u"评论[")[1]
				str3 = str3.split("]")[0]
				cmmt = str3

				str4 = content.split(u" 收藏 ")[-1]
				flag = str4.find(u"来自")
				shijian = str4[:flag]

				weibo_content = content[:content.rindex(u" 赞")]

				#print user_name+temp_leaf_id+temp_weibo_id
				#L000110001 / L000110001 + 0002
				temp_weibo_id = str(weibo_num)
				temp_weibo_id = temp_weibo_id.zfill(4)

				sheet_info.write(temp_row,1,innner_id+temp_weibo_id)				#第01列保存编号
				sheet_info.write(temp_row,2,uid)															#第02列保存用户名
				sheet_info.write(temp_row,3,info_id)														#第03列保存用户地址
				sheet_info.write(temp_row,4,sex)															#第04列保存性别
				sheet_info.write(temp_row,5,loc)															#第05保存地址
				sheet_info.write(temp_row,6,birth)															#第06列保存生日
				sheet_info.write(temp_row,7,quali_id)														#第07列保存认证信息
				sheet_info.write(temp_row,8,str(num_wb))													#第08列保存微博数
				sheet_info.write(temp_row,9,str(num_gz))													#第09列保存关注数
				sheet_info.write(temp_row,10,str(num_fs))													#第10列保存粉丝数
				sheet_info.write(temp_row,11,shijian)															#第11列保存时间戳
				sheet_info.write(temp_row,12,clearAT(weibo_content))													#第12列保存微博内容
				sheet_info.write(temp_row,13,likes)															#第13列保存点赞数
				sheet_info.write(temp_row,14,trans)															#第14列保存转发数
				sheet_info.write(temp_row,15,cmmt)															#第15列保存评论数

				weibo_num += 1
				temp_row += 1

				if weibo_num > 300:
					break;

			page_num += 1
			# https://weibo.cn/u/2714280233?filter=1&page=4
			original_weibo = "https://weibo.cn/u/" + str(info_id) + "?filter=1&page=" + str(page_num)
			#https://weibo.cn/5623115983?filter=1&page=3
			driver.get(original_weibo)
			info = driver.find_elements_by_xpath("//div[@class='c']")
			#if  info.size() == 0:
			#	break

	except Exception,e:
		print "Error: ",e
	finally:
		print u'VisitPersonPage!'
		print '**********************************************\n'

if __name__ == '__main__':
	 #定义变量
    username = '******'             #输入你的用户名
    password = '******'               #输入你的密码
    
    #操作函数
    LoginWeibo(username, password)      #登陆微博

    print 'Read file:'
    workbook = xlrd.open_workbook(user_name + u'.xlsx')
    readSheet = workbook.sheet_by_index(1)

    weibo_id = 1
    weibo_all_num = readSheet.ncols

    weibo_url_col = readSheet.col_values(2)
    for i, cmmt_usr_id in enumerate(weibo_url_col):
    	print u'第'+str(i)+ '/' + str(readSheet.nrows) + u'用户信息'
    	if 'https://weibo.cn/comment/hot' in cmmt_usr_id:
    			continue

    	innner_id = readSheet.row(i)[0]
    	VisitLeafPage(str(cmmt_usr_id),str(innner_id))

    writebook.close()