#!/usr/bin/env python
# coding: utf-8

# In[1]:


from selenium import webdriver
import time
from lxml import etree

def dosth(dw="浙江经贸职业技术学院",qq="1362102080@qq.com"):
    driver=webdriver.Chrome()
    url = 'https://passport.jd.com/uc/login?ltype=logout&ReturnUrl=https://home.jd.com/'
    driver.get(url)
    time.sleep(10)

    driver.find_element_by_xpath('//*[@id="_MYJD_ordercenter"]/a').click()#跳转到我的订单页面
    time.sleep(1)

    driver.find_element_by_xpath('//*[@id="_MYJD_wdfp"]/a').click()#跳转到我的发票页面
    time.sleep(2)


    html = etree.HTML(driver.page_source, etree.HTMLParser())
    infos = html.xpath('//*[@id="main"]/div/div[2]/div[3]/table/tbody')
    for info in infos:
        state = info.xpath('tr[3]/td[3]/span/text()')[0].strip()
        if state == '未开票':
            driver.get("https://jdcs.jd.com/index.action?venderId=106816")
            time.sleep(2)
            driver.find_element_by_xpath("//pre[@name='sendBox']").click()#聚焦
            time.sleep(2)
            driver.find_element_by_xpath("//pre[@name='sendBox']").send_keys("我要开票")#输入文字
            time.sleep(2)
            driver.find_element_by_xpath("//div[@id='app']/div/div/div[2]/div[2]/div/div[2]/div/div[3]/div/div").click()#点击发送
            time.sleep(2)
            driver.find_element_by_xpath("//pre[@name='sendBox']").click()#聚焦
            time.sleep(2)
            driver.find_element_by_xpath("//pre[@name='sendBox']").send_keys(str(dw)+","+str(qq))#输入文字
            time.sleep(2)
            driver.find_element_by_xpath("//div[@id='app']/div/div/div[2]/div[2]/div/div[2]/div/div[3]/div/div").click()#点击发送
            time.sleep(2)
            driver.back()
            time.sleep(2)
        else:
            continue

import datetime
def main(h,m):
    while True:
        now = datetime.datetime.now()
        if now.hour == h and now.minute == m:
            dosth()
        # 每隔60秒检测一次
main(8,0)


# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:




