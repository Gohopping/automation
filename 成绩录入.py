#!/usr/bin/env python
# coding: utf-8

# In[1]:


from selenium import webdriver
from time import sleep
import requests
from lxml import etree
from PIL import Image
from hashlib import md5

driver = webdriver.Chrome()
driver.maximize_window()

driver.get('http://jxgl.zjiet.edu.cn/(S(eai4jbq3uruazza3llmp0rub))/js_main.aspx?xh=2165')


# In[2]:


#标签定位
search_input = driver.find_element_by_id('TextBox1')
#标签交互
search_input.send_keys('xxxxxxxxxxx')


# In[3]:


#标签定位
search_input = driver.find_element_by_id('TextBox2')
#标签交互
search_input.send_keys('xxxxxxxxxxx')


# In[4]:


driver.save_screenshot('a.png')
imgelement = driver.find_element_by_xpath('//*[@id="icode"]')   #定位标签
location = imgelement.location
size = imgelement.size
rangle = (int(location['x']), int(location['y']), int(location['x'] + size['width']),
          int(location['y'] + size['height']))  # 写成我们需要截取的位置坐标
i = Image.open("a.png")  # 打开截图
frame4 = i.crop(rangle)  # 使用Image的crop函数，从截图中再次截取我们需要的区域
frame4.save('save.png') # 保存我们接下来的验证码图片 进行打码


# In[5]:


class Chaojiying_Client(object):

    def __init__(self, username, password, soft_id):
        self.username = username
        password =  password.encode('utf8')
        self.password = md5(password).hexdigest()
        self.soft_id = soft_id
        self.base_params = {
            'user': self.username,
            'pass2': self.password,
            'softid': self.soft_id,
        }
        self.headers = {
            'Connection': 'Keep-Alive',
            'User-Agent': 'Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 5.1; Trident/4.0)',
        }

    def PostPic(self, im, codetype):
        """
        im: 图片字节
        codetype: 题目类型 参考 http://www.chaojiying.com/price.html
        """
        params = {
            'codetype': codetype,
        }
        params.update(self.base_params)
        files = {'userfile': ('ccc.jpg', im)}
        result = requests.post('http://upload.chaojiying.net/Upload/Processing.php', data=params, files=files, headers=self.headers).json()
        return result

    def ReportError(self, im_id):
        """
        im_id:报错题目的图片ID
        """
        params = {
            'id': im_id,
        }
        params.update(self.base_params)
        r = requests.post('http://upload.chaojiying.net/Upload/ReportError.php', data=params, headers=self.headers)
        return r.json()



chaojiying = Chaojiying_Client('1551969060', 'luoli860826', '909607')	#用户中心>>软件ID 生成一个替换 96001
im = open('save.png', 'rb').read()													#本地图片文件路径 来替换 a.jpg 有时WIN系统须要//
result = chaojiying.PostPic(im, 1004)
print(result)
print('验证码结果为:' + result['pic_str'])


# In[6]:


#标签定位
search_input = driver.find_element_by_id('TextBox3')
#标签交互
search_input.send_keys(result['pic_str'])


# In[7]:


btn = driver.find_element_by_id('Button1')
btn.click()


# In[8]:


import xlrd
import xlwt
#打开文件
workbook = xlrd.open_workbook("D:\成绩表.xlsx")
Sheet = workbook.sheet_by_name('Sheet1')
xh_list = []
xm_list = []
km_list = []
cj_list = []
#任课信息
rkls = Sheet.cell_value(1,5)
print(rkls)
#任课老师
mlz = Sheet.cell_value(2,5)
print(mlz)

for xh in range(42):
    xh_list.append(Sheet.cell_value(xh,0))
print (xh_list)
for xm in range(42):
    xm_list.append(Sheet.cell_value(xm,1))
print(xm_list)
for km in range(42):
    km_list.append(Sheet.cell_value(km,2))
print(km_list)
for cj in range(42):
    cj_list.append(Sheet.cell_value(cj,3))
print(cj_list)


# In[ ]:




