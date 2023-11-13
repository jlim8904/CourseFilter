from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from getpass import getpass
import pandas as pd

# 登入
def login():
    print("登入中...")
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="iosfix"]/md-toolbar/div/button')))
    btn_first = driver.find_element(By.XPATH, '//*[@id="iosfix"]/md-toolbar/div/button')
    btn_first.click()
    
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, 'input_117')))
    name_ele = driver.find_element(By.XPATH,'//*[@id="input_117"]')
    name_ele.send_keys(account_name)

    password_ele = driver.find_element(By.XPATH, '//*[@id="input_118"]')
    password_ele.send_keys(account_password)

    btn_last = driver.find_elements(By.XPATH, '/html/body/div[3]/md-dialog/form/md-dialog-actions/button[2]')
    btn_last[0].click()
    print("登入成功!")
    
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.CLASS_NAME, "timetable.ng-scope.flex-grow")))
    timetable = driver.find_element(By.CLASS_NAME, 'timetable.ng-scope.flex-grow')
    nameEle = timetable.find_element(By.CLASS_NAME, 'ng-binding').text
    nameSplit = nameEle.split("：")
    return nameSplit[-1]
    

def selected_schedule():
    # 取得已選課程
    print("正在讀取已選課程資料...")
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.CLASS_NAME, "timetable-item.ng-scope.layout-align-center-stretch.layout-column.flex")))
    sel_class = driver.find_elements(By.CLASS_NAME, 'timetable-item.ng-scope.sel.layout-align-center-stretch.layout-column.flex')
    sel_class_list = []

    for sel in sel_class:
        sel_class_list.append(sel.text)
    sel_class_list = list(set(sel_class_list))
    
    # 查詢已選課程
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.CLASS_NAME, "md-tab.ng-scope.ng-isolate-scope")))
    conditionsearchBtn = driver.find_elements(By.CLASS_NAME, "md-tab.ng-scope.ng-isolate-scope")
    conditionsearchBtn[1].click()
    
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, 'input_44')))
    input_box = driver.find_element(By.ID, "input_44")
    for c in sel_class_list:
        # 填入選課代號
        input_box.clear()
        input_box.send_keys(c)
        # 點擊查詢
        searchBtn = driver.find_elements(By.CLASS_NAME, "md-raised.md-primary.md-button.ng-scope")
        searchBtn[1].click()
        
        # 爬取課程
        WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "//*[@class='rwd-table mid ng-scope']/tbody")))
        class_data_table = driver.find_element(By.XPATH, "//*[@class='rwd-table mid ng-scope']/tbody")
        class_data = class_data_table.find_elements(By.CLASS_NAME, "ng-scope")
        for d in class_data:
            for data in d.find_elements(By.CSS_SELECTOR, "*"):
                if data.get_attribute('data-title') == "上課時間/上課教室/授課教師":
                    course_details = data.text
                    for details in course_details.split():
                        if details[0] == '(':
                            lesson = details.split(')')[1]
                            if '-' in lesson:
                                start = int(lesson.split('-')[0])
                                end = int(lesson.split('-')[1])
                            else:
                                start = end = int(lesson)
                            for n, day in enumerate(days):
                                if day == details[1]:
                                    for i in range(start-1, end):
                                        schedule[n][i] = 1
                                    break
                    break
                break
        # 重新查詢
        WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.CLASS_NAME, "md-raised.md-warn.md-button.ng-scope")))
        searchBtn = driver.find_element(By.CLASS_NAME, "md-raised.md-warn.md-button.ng-scope")
        searchBtn.click()


def search_dept_course():
    dept_option = list(range(120,138))
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.CLASS_NAME, "md-tab.ng-scope.ng-isolate-scope")))
    deptsearchBtn = driver.find_element(By.CLASS_NAME, "md-tab.ng-scope.ng-isolate-scope")
    deptsearchBtn.click()

    deptFlag = 0
    deptCount = 0
    for option in dept_option:
        # 點擊學院列表 
        if deptFlag == 0:
            WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, 'select_34')))
            deptBtn = driver.find_element(By.ID, 'select_34')
            deptBtn.click()
            
        WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, f'select_option_{option}')))
        dept = driver.execute_script(f'return document.getElementById("select_option_{option}");').text
        
        if not all_dept and dept not in select_dept:
            deptFlag = 1
            continue
        deptFlag = 0
        deptCount += 1
        
        # 點擊學院
        driver.execute_script(f'return document.getElementById("select_option_{option}").click();')
        
        # 點擊查詢
        WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.CLASS_NAME, "md-raised.md-primary.md-button.ng-scope")))
        searchBtn = driver.find_element(By.CLASS_NAME, "md-raised.md-primary.md-button.ng-scope")
        searchBtn.click()
        
        WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.CLASS_NAME, "ng-scope.layout-align-space-between-center.layout-row")))
        totalClassDiv = driver.find_element(By.CLASS_NAME, "ng-scope.layout-align-space-between-center.layout-row")
        totalClass = int(totalClassDiv.find_element(By.CLASS_NAME, "ng-binding").text)
        loadTimes = int((totalClass + 1) / 50)
        
        # 載入更多
        for i in range(loadTimes):
            WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.CLASS_NAME, "md-raised.md-primary.md-hue.md-button.ng-scope")))
            loadBtn = driver.find_element(By.CLASS_NAME, "md-raised.md-primary.md-hue.md-button.ng-scope")
            loadBtn.click()
                
        print(f"正在讀取{dept}頁面...")
        # 爬取課程
        WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "//*[@class='rwd-table mid ng-scope']/tbody")))
        class_data_table = driver.find_element(By.XPATH, "//*[@class='rwd-table mid ng-scope']/tbody")
        class_data = class_data_table.text.split()
        i = 0
        while i < len(class_data):
            if (class_data[i] == '關注' or class_data[i] == '取消') and class_data[i+1].isnumeric():
                cid = class_data[i+1]
                i+=3
                all_course_data[cid] = {'課程名稱': class_data[i]}
                while not ('/' in class_data[i] and class_data[i].split('/')[0].isnumeric() and class_data[i].split('/')[1].isnumeric()):
                    i+=1
                    if class_data[i][0] == '(':
                        try:
                            all_course_data[cid]['上課時間'] += f' {class_data[i]}'
                        except:
                            all_course_data[cid]['上課時間'] = f'{class_data[i]}'
                if not any([ele.isdigit() for ele in class_data[i-1]]) and not class_data[i-1] == '未排教室':
                    all_course_data[cid]['授課教師'] = class_data[i-1]
                quota = class_data[i].split('/')
                all_course_data[cid]['剩餘人數'] = int(quota[1]) - int(quota[0])
            i+=1
        # 重新查詢
        WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.CLASS_NAME, "md-raised.md-warn.md-button.ng-scope")))
        searchBtn = driver.find_element(By.CLASS_NAME, "md-raised.md-warn.md-button.ng-scope")
        searchBtn.click()
        if not all_dept and len(select_dept) == deptCount:
            break


def get_course_not_conflict():
    print()
    count = 0
    for course in all_course_data:
        count+=1
        print(f"\033[1A\033[K未衝堂課程篩選中...\t{round((count/len(all_course_data))*100)}%")
        conflict = 0
        course_details = all_course_data[course]['上課時間']
        for details in course_details.split():
            if details[0] == '(':
                lesson = details.split(')')[1]
                if '-' in lesson:
                    start = int(lesson.split('-')[0])
                    end = int(lesson.split('-')[1])
                else:
                    try:
                        start = end = int(lesson)
                    except:
                        start = 1
                        end = 0
                for n, day in enumerate(days):
                    if day == details[1]:
                        for i in range(start-1, end):
                            if schedule[n][i] == 1:
                                conflict = 1
                                break
                        break
        # 印出不衝堂且剩餘人數大於0的課程
        if not conflict and all_course_data[course]['剩餘人數'] > 0:
            course_not_conflict[course] = all_course_data[course]
    for course in course_not_conflict:
        print(course, course_not_conflict[course])
    print(f'共計{len(course_not_conflict)}門未衝堂課程')


def course_to_excel():
    print(f'正在輸出EXCEL檔...')
    course_not_conflict_df = pd.DataFrame(course_not_conflict).T
    course_not_conflict_df.to_excel(f"未衝堂課程_{name}{account_name}.xlsx", index_label='課程代碼')
    print(f'已輸出EXCEL檔')


def dcard():
    print(f'網頁:{driver.title}')
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);") 
        
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, "atm_40_ncl75p.atm_gi_1d1uzc4.atm_mk_h2mmj6.atm_lo_1lq3voq.atm_le_1lq3voq.atm_lk_18fnu1o.atm_ll_18fnu1o.atm_9s_11p5wf0.atm_dy_m63k86.atm_dz_1ghlemp.atm_9j_tlke0l.atm_r2_1j28jx2.atm_e0_jok701.c122gkvw")))
    ele = driver.find_elements(By.CLASS_NAME, "atm_40_ncl75p.atm_gi_1d1uzc4.atm_mk_h2mmj6.atm_lo_1lq3voq.atm_le_1lq3voq.atm_lk_18fnu1o.atm_ll_18fnu1o.atm_9s_11p5wf0.atm_dy_m63k86.atm_dz_1ghlemp.atm_9j_tlke0l.atm_r2_1j28jx2.atm_e0_jok701.c122gkvw")
    print(len(ele))
    count = 1

    for i in range(len(ele)):
        title = ele[i].find_elements(By.CLASS_NAME, "atm_cs_1urozh.atm_c8_1csq7v7.atm_g3_1qqjw7d.atm_7l_1pday2.atm_1938jqx_1yyfdc7.atm_2zt8x3_stnw88.atm_grwvqw_gknzbh.atm_1ymp90q_idpfg4.atm_89ifzh_idpfg4.atm_1hh4tvs_1osqo2v.atm_1054lsl_1osqo2v.t1gihpsa")
        a_ele = ele[i].find_element(By.CSS_SELECTOR, "a")
        link = a_ele.get_attribute('href')
        if (len(title) != 0):
            print(f'{count}: {title[0].text}')
            print(link)  # 取得文章連結
            count = count+1


if __name__ == '__main__':
    account_name = input("Account: ")
    account_password = getpass("Password: ")
    
    url =   "https://coursesearch01.fcu.edu.tw"
    
    days = ['一','二','三','四','五','六','日']
    schedule = [[0]*14 for _ in range(7)]
    select_dept = ['創能學院','資電學院','外語文','通識核心課']
    all_dept = False
    
    all_course_data = {}
    course_not_conflict = {}
    
    driver = webdriver.Chrome(service=Service("chromedriver.exe"))
    
    driver.get(url)

    name = login()
    search_dept_course()
    selected_schedule()
    get_course_not_conflict()
    course_to_excel()

    # Dcard search
    while True:
        courseNum = input("請輸入欲搜尋的課程代碼: ")
        if courseNum == '-1':
            break
        try:
            cid = all_course_data[courseNum]['課程名稱']
        except:
            print("課程代碼輸入錯誤，請重新輸入")
            continue
        cname = all_course_data[courseNum]['授課教師'] if '授課教師' in all_course_data[courseNum] else ''
        dcardURL = f"https://www.dcard.tw/search?forum=fcu&query={cid}%20{cname}"
        driver = webdriver.Chrome(executable_path="chromedriver.exe")
        driver.get(dcardURL)  # 通過瀏覽器開啟url
        dcard()
    driver.close()
