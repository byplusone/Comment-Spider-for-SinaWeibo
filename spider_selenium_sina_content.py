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
# 垃圾xlwt，最多只能操作256的数组大小
# import xlwt
# 垃圾openpyxl一点算数不友好
# from openpyxl.workbook import Workbook  
#万恶的ExcelWriter，妹的封装好了不早说，封装了很强大的excel写的功能  
# from openpyxl.writer.excel import ExcelWriter  
#一个eggache的数字转为列字母的方法  
# from openpyxl.cell import get_column_letter
import xlsxwriter


#先调用无界面浏览器PhantomJS或Firefox    
#driver = webdriver.PhantomJS(executable_path="G:\phantomjs-1.9.1-windows\phantomjs.exe")    
driver = webdriver.Firefox()
#options = webdriver.ChromeOptions()
#options.add_experimental_option("excludeSwitches", ["ignore-certificate-errors"])
#driver = webdriver.Chrome(chrome_options=options)
wait = ui.WebDriverWait(driver,10)


#全局变量 文件操作读写信息
#inforead = codecs.open("SinaWeibo_List.txt", 'r', 'utf-8')
writebook = xlsxwriter.Workbook('Expenses02.xlsx')
sheet_info = writebook.add_worksheet()
comment_user_info = writebook.add_worksheet()
infofile = codecs.open("SinaWeibo_Info1.txt", 'a', 'utf-8')
temp_row = 1 # 定义写操作时全局的当前操作的行数
#infofile =  xlwt.Workbook()


#********************************************************************************
#                  第一步: 登陆weibo.cn 获取新浪微博的cookie
#        该方法针对weibo.cn有效(明文形式传输数据) weibo.com见学弟设置POST和Header方法
#                LoginWeibo(username, password) 参数用户名 密码

#        https://www.zhihu.com/question/21451510
#        http://www.cnblogs.com/fnng/p/3606934.html
#                             验证码暂停时间手动输入
#********************************************************************************

def LoginWeibo(username, password):
    try:
        #**********************************************************************
        # 直接访问driver.get("http://weibo.cn/5824697471")会跳转到登陆页面 用户id
        #
        # 用户名<input name="mobile" size="30" value="" type="text"></input>
        # 密码 "password_4903" 中数字会变动,故采用绝对路径方法,否则不能定位到元素
        #
        # 勾选记住登录状态check默认是保留 故注释掉该代码 不保留Cookie 则'expiry'=None
        #**********************************************************************
        
        #输入用户名/密码登录
        print u'准备登陆Weibo.cn网站...'
        driver.get("http://login.weibo.cn/login/")
        elem_user = driver.find_element_by_name("mobile")
        elem_user.send_keys(username) #用户名
        elem_pwd = driver.find_element_by_xpath("/html/body/div[2]/form/div/input[2]")
        elem_pwd.send_keys(password)  #密码
        #elem_rem = driver.find_element_by_name("remember")
        #elem_rem.click()             #记住登录状态

        #重点: 暂停时间输入验证码
        #pause(millisenconds)
        time.sleep(20)
        
        elem_sub = driver.find_element_by_name("submit")
        elem_sub.click()              #点击登陆
        time.sleep(2)
        
        #获取Coockie 推荐 http://www.cnblogs.com/fnng/p/3269450.html
        print driver.current_url
        print driver.get_cookies()  #获得cookie信息 dict存储
        print u'输出Cookie键值对信息:'
        for cookie in driver.get_cookies(): 
            #print cookie
            for key in cookie:
                print key, cookie[key]
                    
        #driver.get_cookies()类型list 仅包含一个元素cookie类型dict
        print u'登陆成功...'
        
        
    except Exception,e:      
        print "Error: ",e
    finally:    
        print u'End LoginWeibo!\n\n'


#********************************************************************************
#                  第二步: 访问个人页面http://weibo.cn/5824697471并获取信息
#                                VisitPersonPage()
#        编码常见错误 UnicodeEncodeError: 'ascii' codec can't encode characters 
#********************************************************************************

def VisitPersonPage(user_id,user_num):

    try:
        global infofile
        global temp_row
        
        print u'准备访问个人网站.....'
        #原创内容 http://weibo.cn/guangxianliuyan?filter=1&page=2
        driver.get("http://weibo.cn/" + user_id)

        #**************************************************************************
        # No.1 直接获取 用户昵称 微博数 关注数 粉丝数
        #      str_name.text是unicode编码类型
        #**************************************************************************

        #用户id
        print u'个人详细信息'
        print '**********************************************'
        print u'用户id: ' + user_id

        #昵称
        str_name = driver.find_element_by_xpath("//div[@class='ut']")
        str_t = str_name.text.split(" ")
        num_name = str_t[0]      #空格分隔 获取第一个值 "Eastmount 详细资料 设置 新手区"
        print u'昵称: ' + num_name 

        #微博数 除个人主页 它默认直接显示微博数 无超链接
        #Error:  'unicode' object is not callable
        #一般是把字符串当做函数使用了 str定义成字符串 而str()函数再次使用时报错
        str_wb = driver.find_element_by_xpath("//div[@class='tip2']")  
        pattern = r"\d+\.?\d*"   #正则提取"微博[0]" 但r"(\[.*?\])"总含[] 
        guid = re.findall(pattern, str_wb.text, re.S|re.M)
        print str_wb.text        #微博[294] 关注[351] 粉丝[294] 分组[1] @他的
        for value in guid:
            num_wb = int(value)
            break
        print u'微博数: ' + str(num_wb)

        #关注数
        str_gz = driver.find_element_by_xpath("//div[@class='tip2']/a[1]")
        guid = re.findall(pattern, str_gz.text, re.M)
        num_gz = int(guid[0])
        print u'关注数: ' + str(num_gz)

        #粉丝数
        str_fs = driver.find_element_by_xpath("//div[@class='tip2']/a[2]")
        guid = re.findall(pattern, str_fs.text, re.M)
        num_fs = int(guid[0])
        print u'粉丝数: ' + str(num_fs)
        

        #***************************************************************************
        # No.2 文件操作写入信息
        #***************************************************************************

        infofile.write('=====================================================================\r\n')
        infofile.write(u'用户: ' + user_id + '\r\n')
        infofile.write(u'昵称: ' + num_name + '\r\n')
        infofile.write(u'微博数: ' + str(num_wb) + '\r\n')
        infofile.write(u'关注数: ' + str(num_gz) + '\r\n')
        infofile.write(u'粉丝数: ' + str(num_fs) + '\r\n')
        infofile.write(u'微博内容: ' + '\r\n\r\n')
        
        

        #***************************************************************************
        # No.3 获取关注人列表
        # http://weibo.cn/guangxianliuyan?filter=0&page=1
        # 其中filter=0表示全部 =1表示原创
        # weibo.cn/comment/EcXMmsS1o?uid=3261134763&rl=0&page=3
        # weibo.cn/comment/EcXMmsS1o?uid=3261134763&rl=0&page=2
        # weibo.cn/comment/EcXMmsS1o?uid=3261134763&rl=0#cmtfrm
        #***************************************************************************

        print '\n'
        print u'获取微博内容信息'
        num = 1
        #共查找1页的内容
        comment_num = []
        while num <= 1:
        	#对原创微博进行爬取，内容质量较为正常，并且不会出现原微博评论链接混淆的情况
            url_wb = "http://weibo.cn/" + user_id + "?filter=1&page=" + str(num)
            print url_wb
            driver.get(url_wb)
            info = driver.find_elements_by_xpath("//div[@class='c']")
            #print info #这里info是一组数，一个value代表一条消息
            #link_location =  driver.find_elements_by_xpath("//div[@class='c']/div[@class='cc']")
            #获取评论连接位置
            weibo_num = 0
            for value in info:
                #print value.text
                content = value.text
                #print content

                #跳过最后一行数据为class=c
                #Error:  'NoneType' object has no attribute 'groups'
                if u'设置:皮肤.图片' not in content:
                    if content.startswith(u'转发'):
                        print u'转发微博'
                        infofile.write(u'转发微博\r\n')
                    else:
                        print u'原创微博'
                        infofile.write(u'原创微博\r\n')
                        
                    #获取最后一个点赞数 因为转发是后有个点赞数
                    str1 = content.split(u" 赞")[-1]
                    if str1: 
                        val1 = re.match(r'\[(.*?)\]', str1).groups()[0]
                        print u'点赞数: ' + val1
                        infofile.write(u'点赞数: ' + str(val1) + '\r\n')

                    str2 = content.split(u" 转发")[-1]
                    if str2: 
                        val2 = re.match(r'\[(.*?)\]', str2).groups()[0]
                        print u'转发数: ' + val2
                        infofile.write(u'转发数: ' + str(val2) + '\r\n')

                    str3 = content.split(u" 评论")[-1]
                    if str3:
                        val3 = re.match(r'\[(.*?)\]', str3).groups()[0]
                        #用队列comment_num储存每条微博的评论数
                        comment_num.append(int(val3))
                        print u'评论数: ' + val3
                        infofile.write(u'评论数: ' + str(val3) + '\r\n')

                    str4 = content.split(u" 收藏 ")[-1]
                    flag = str4.find(u"来自")
                    print u'时间: ' + str4[:flag]
                    infofile.write(u'时间: ' + str4[:flag] + '\r\n')

                    print u'微博内容:'
                    print content[:content.rindex(u" 赞")]  #后去最后一个赞位置
                    infofile.write(content[:content.rindex(u" 赞")] + '\r\n')
                    infofile.write('\r\n')
                    print '\n'
    
                    sheet_info.write(temp_row + weibo_num,1,user_num)							#第一列保存编号
                    sheet_info.write(temp_row + weibo_num,2,num_name)							#第二列保存昵称
                    sheet_info.write(temp_row + weibo_num,3,user_id) 							#第三列保存id
                    sheet_info.write(temp_row + weibo_num,4,str4[:flag]) 						#第四列保存时间
                    sheet_info.write(temp_row + weibo_num,5,content[:content.rindex(u" 赞")]) 	#第五列保存内容
                    sheet_info.write(temp_row + weibo_num,6,str(val1))							#第六列保存点赞数
                    sheet_info.write(temp_row + weibo_num,7,str(val2))							#第七列保存转发数
                    sheet_info.write(temp_row + weibo_num,8,str(val3))							#第八列保存评论数
					
                    weibo_num += 1

                else:
                    print u'跳过', content, '\n'
                    break
            else:
                print u'next page...\n'
                infofile.write('\r\n\r\n')

            # 以下为手机端评论URL样例
            # weibo.cn/comment/EcXMmsS1o?uid=3261134763&rl=0&page=3
            # weibo.cn/comment/EcXMmsS1o?uid=3261134763&rl=0&page=2
            # weibo.cn/comment/EcXMmsS1o?uid=3261134763&rl=0#cmtfrm
            # 一页评论显示10条评论，这里变量cmmt_page*10就代表了评论的数量
            # comment存储了所有评论的位置
            # 这里报错Stale Element Reference Exception，所以决定第一次将一个用户对应微博下的所有评论URL保存在list中
            comment = driver.find_elements_by_xpath("//a[@class='cc']")
            comment_URL_list = []
            for url_cmmt in comment:
                comment_URL = url_cmmt.get_attribute("href")
                comment_URL = comment_URL.split('#')[0]
                comment_URL_list.append(comment_URL)

            weibo_num = -1
            for url_cmmt in comment:
            	weibo_num += 1
            	i = 0 #评论计数
            	temp_comment_num = 0 #评论用户地址计数
            	temp_page = comment_URL_list.pop(0)
            	cmmt_page = 1
            	temp_cmmt_num = comment_num.pop(0)
            	# 无法判断是否是热评，只能选择对于评论数过多的微博直接减去一个预测的首页热评数
            	# 新浪微博显示的评论数有水分，统计实验一条显示评论为270的微博实际可以看到的评论为242，一条显示评论为1670的微博实际只可以看到443
            	if temp_cmmt_num > 20:
            		temp_cmmt_num -= 15
            	#elif temp_cmmt_num > 100 and temp_cmmt_num <=400:
            	#	temp_cmmt_num -= 40 # 发现400条以内的微博40条评论水分已经不算多了2333

            	while i < 400 and temp_cmmt_num > 0:
                	dest_URL = temp_page+ "&page=" + str(cmmt_page)
                	driver.get(dest_URL)
                	if cmmt_page == 1:
                		cmmt_text = driver.find_elements_by_xpath
                	cmmt_text = driver.find_elements_by_xpath("//div[@class='c']/span[@class='ctt']")
                	# 微博评论的第一页会出现热门评论，导致最后temp_cmmt_num无法归零
                	# 已通过直接减除水分解决
                	if cmmt_text:
                		print '**********************************************\n'
                		for j,value in enumerate(cmmt_text):   	        
                			info = value.text
                			sheet_info.write(temp_row + weibo_num,i+9,info)
                			i += 1
                			temp_cmmt_num -= 1
                			print info
                			print '\n'
                			print temp_cmmt_num
                			print '\n'
                		print '**********************************************\n'
                		cmmt_user = driver.find_elements_by_xpath("//div[@class='c']/a[1]")
                		for p,value in enumerate(cmmt_user):
                			value = value.get_attribute("href")
                			if p == 0 or p == len(cmmt_user)-1:
                				continue
                			else:
                				comment_user_info.write(temp_comment_num,temp_row + weibo_num, value)
                				temp_comment_num += 1
                				print value
        	        	cmmt_page += 1
        	        else:
        	        	print "2333333333333333333333"
        	        	break
            num += 1
            temp_row += weibo_num
            print '\n\n'
        print '**********************************************'
            
        
        
    except Exception,e:      
        print "Error: ",e
    finally:    
        print u'VisitPersonPage!\n\n'
        print '**********************************************\n'
        

    
#*******************************************************************************
#                                程序入口 预先调用
#*******************************************************************************
    
if __name__ == '__main__':

    #定义变量
    username = '******'             #输入你的用户名
    password = '******'               #输入你的密码
    
    #用户id url+id访问个人
    #user_id = 'renzhiqiang'
    #user_id = 'guangxianliuyan'  

    

    #操作函数
    LoginWeibo(username, password)      #登陆微博

    #driver.add_cookie({'name':'name', 'value':'_T_WM'})
    #driver.add_cookie({'name':'value', 'value':'c86fbdcd26505c256a1504b9273df8ba'})

    #注意
    #因为sina微博增加了验证码,但是你用Firefox登陆一次输入验证码,再调用该程序即可,因为Cookies已经保证
    #会直接跳转到明星微博那部分,即: http://weibo.cn/guangxianliuyan
    

    #在if __name__ == '__main__':引用全局变量不需要定义 global inforead 省略即可
    print 'Read file:'
    #user_id = inforead.readline()
    workbook = xlrd.open_workbook(u"S_test3.xlsx")
    #while user_id!="":
        #user_id = user_id.rstrip('\r\n')
        #VisitPersonPage(user_id)         #访问个人页面
        #user_id = inforead.readline()
        #break
    sheet1 = workbook.sheet_by_index(0)
    cols = sheet1.col_values(3) # 获取第四列内容
    cols_user_num = sheet1.col_values(0) # 获取第一列的用户编号
    for num,user_id in enumerate(cols):
        print user_id
        user_num = cols_user_num[num - 1]# num从1开始，访问数组则应从0开始
        VisitPersonPage(str(int(user_id)),user_num)

    writebook.close()
    infofile.close()
    #inforead.close()
    
    

    
