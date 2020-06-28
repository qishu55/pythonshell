from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import xlrd

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
shanchu = []
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
        time.sleep(1.5)
        #不需要禁用则去掉
        # jinyong = driver.find_elements_by_class_name("btn_gray")[2].text
        # if jinyong == "启用":
        #     print(lizhi+"已经是禁用状态")
        #     driver.find_element_by_class_name("btn_select_txt").click()
        #     return
        # else:
            #删除操作
        driver.find_elements_by_class_name("btn_gray")[3].click()
        time.sleep(1)
        driver.find_element_by_class_name("btnSubmit").click()
        time.sleep(1)
        driver.find_element_by_class_name("qy_dlgInput").send_keys("backup@lightinthebox.com")
        time.sleep(1)
        driver.find_element_by_class_name("o_title").click()
        time.sleep(1.5)
        driver.find_element_by_class_name("btnSubmit").click()
        print(lizhi+'账号已删除')
        shanchu.append(lizhi)
        return shanchu
    except:
        #没有则提示不存在
        print(lizhi+':该用户不存在')
    return shanchu


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

print(shanchu)