from selenium import webdriver
import time
import xlrd
#删除太多会有验证码提示
driver = webdriver.Firefox()
# 访问百度首页
first_url= 'https://exmail.qq.com/login'
print("now access %s" %(first_url))
driver.get(first_url)




# driver.find_element_by_id("kw").clear()
driver.find_element_by_class_name("js_show_pwd_panel").click()
driver.find_element_by_id("inputuin").send_keys("luozhiyu@lightinthebox.com")
driver.find_element_by_id("pp").send_keys("qishu123A")
driver.find_element_by_id("btlogin").click()

def serach(lizhi):
    #查找人员信息
    # search_text = driver.find_element_by_name("dy_content")
    time.sleep(3)
    search_text = driver.find_element_by_name("dy_content")
    search_text.clear()
    search_text.send_keys(lizhi)
    driver.find_element_by_class_name("btnSearch").click()

    time.sleep(1.5)
    try:
        #找到了账号则禁用
        search_text2 = driver.find_element_by_link_text(lizhi)
        search_text2.click()
        time.sleep(2.5)
        jinyong = driver.find_elements_by_class_name("btn_gray")[2].text
        if jinyong == "启用":
            print(lizhi+"已经是禁用状态")
            driver.find_element_by_class_name("btn_select_txt").click()
            return
        else:
            #禁用操作
            driver.find_elements_by_class_name("btn_gray")[2].click()
            driver.find_element_by_class_name("btnSubmit").click()
            driver.find_element_by_class_name("btn_select_txt").click()
            print(lizhi+'已禁用')
            return
    except:
        #没有则提示不存在
        print(lizhi+':该用户不存在')
    return


for i in range(10):
    print("等待的第{}秒".format(i))
    time.sleep(1)
driver.find_element_by_id("manageMail").click()
driver.find_element_by_link_text("成员与分组").click()
# 输入离职账号查询
x1 = xlrd.open_workbook("1.xlsx")
sheet1 = x1.sheet_by_name("Sheet1")
list1 = sheet1.nrows
for i in range(list1):
    lizhi = sheet1.row_values(i)[0]

    serach(lizhi)
    print("-"*20)