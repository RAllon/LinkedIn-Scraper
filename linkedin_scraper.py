import parameters
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from parsel import Selector
import xlwt
import xlrd

loc =('path to your excel file with names and institutions')
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

style0 = xlwt.easyxf('font: name Times New Roman, color-index black, bold on', num_format_str='#,##0.00')
wd = xlwt.Workbook()
ws = wd.add_sheet('results_file_6')

ws.write(0, 0, "Name", style0)
ws.write(0, 1, "College", style0)
ws.write(0, 2, "URL", style0)

name_list = []
prefered_list = []
college_list = []
graduate_list = []
search_query_list = []
new_linkedin_urls = []

driver = webdriver.Chrome('/Users/path to chromedriver/chromedriver')
driver.get('https://www.linkedin.com/login?fromSignIn=true&trk=guest_homepage-basic_nav-header-signin')

username = driver.find_element_by_id('username')
username.send_keys('your linkedin username')
sleep(0.5)

password = driver.find_element_by_id('password')
password.send_keys('your linkedin password')
sleep(0.5)

log_in_button = driver.find_element_by_class_name('login__form_action_container')
log_in_button.click()
sleep(0.5)

driver.get('https://www.google.com/')
sleep(3)

x = 1
while x < 8300: #number of rows
    first_name = sheet.cell_value(x, 1)
    last_name = sheet.cell_value(x, 3)
    full = "{} {}".format(first_name, last_name)
    name_list.append(full)

    prefered_name = sheet.cell_value(x, 2)
    if prefered_name == 0 or prefered_name == first_name:
        prefered_full = "null"
    else:
        prefered_full = "{} {}".format(prefered_name, last_name)
    prefered_list.append(prefered_full)

    college_name = sheet.cell_value(x,4)
    college_list.append(sheet.cell_value(x,4))

    graduate = sheet.cell_value(x, 6)
    if graduate == 0 or college_name == graduate:
        graduate_full = "null"
    else:
        graduate_full = sheet.cell_value(x, 6)
    graduate_list.append(graduate_full)

    #together = 'site:linkedin.com/in/, {}, {}, {}, {}'.format(full, sheet.cell_value(x,4), graduate_full, prefered_full)
    together = 'linkedin {}, {}, {}, {}'.format(full, sheet.cell_value(x,4), graduate_full, prefered_full)
    tog_index = together.find(', null')
    together_new = together[:tog_index] + together[tog_index + 6:]

    tog_index_2 = together_new.find(', null')
    together_new_2 = together_new[:tog_index] + together_new[tog_index + 6:]
    search_query_list.append(together_new_2)

    x = x + 1

z = 0
y=1
while z < 8299:
    search = search_query_list[z]
    print(search)
    driver.get('https://www.google.com/')
    search_query = driver.find_element_by_name('q')
    search_query.send_keys(search)
    sleep(1)
    search_query.send_keys(Keys.RETURN)
    sleep(3)

    linkedin_urls = driver.find_elements_by_class_name('iUh30')
    linkedin_urls = [url.text for url in linkedin_urls]
    sleep(0.5)

    new_linkedin_urls = []
    real_linkedin_urls=[]

    for linkedin_url in linkedin_urls:
        index = linkedin_url.find(" › ")
        linkedin_url = linkedin_url[:index] + '/in/' + linkedin_url[index + 3:]
        new_linkedin_url = 'https://' + linkedin_url
        new_linkedin_urls.append(new_linkedin_url)
    new_linkedin_urls=list(filter(('https:///in/').__ne__,new_linkedin_urls))
    for string in new_linkedin_urls:
        if "linkedin" in string:
            real_linkedin_urls.append(string)

    length = len(real_linkedin_urls)
    if length >= 3:
        k = length - 3
        real_linkedin_urls = real_linkedin_urls[: len(real_linkedin_urls) - k]
    if "https://www.linkedin.com/in/unavailable/" in real_linkedin_urls:
        real_linkedin_urls.remove("https://www.linkedin.com/in/unavailable/")

    for new_linkedin_url in real_linkedin_urls:
        driver.get(new_linkedin_url) # get the profile URL
        sleep(5) # add a 5 second pause loading each URL
        sel = Selector(text=driver.page_source) # assigning the source code for the webpage to variable sel
        Name_Header = driver.find_elements_by_xpath("/html/body/div[7]/div[3]/div/div/div/div/div[2]/main/div[1]/section/div[2]/div[2]/div[1]/ul[1]/li[1]")
        if Name_Header == []:
            Name_Header = driver.find_elements_by_xpath("/html/body/div[8]/div[3]/div/div/div/div/div[2]/main/div[1]/section/div[2]/div[2]/div[1]/ul[1]/li[1]")
        for span in Name_Header:
            name = span.text
        College_Header = driver.find_elements_by_xpath("/html/body/div[7]/div[3]/div/div/div/div/div[2]/main/div[2]/div[5]/span/div/section/div[2]/section/ul/li[2]/div/div/a/div[2]/div/h3")
        if College_Header == []:
            college = "null"
        else:
            for span in College_Header:
                college = span.text
        linkedin_url = driver.current_url
        if linkedin_url == "https://www.linkedin.com/in/unavailable/":
            y = y
        elif name == name_list[z] or name == prefered_list[z] and college == college_list[z]:
            ws.write(y, 0, name)
            ws.write(y, 1, college)
            ws.write(y, 2, linkedin_url)
            wd.save('results_file_6.xls')
            y = y + 1
    z = z + 1

wd.save('results_file_6.xls')
driver.quit()
