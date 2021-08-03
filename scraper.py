
from openpyxl import Workbook
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By

s=Service("C:\\Users\\Sony\\Desktop\\chromedriver.exe")
driver=webdriver.Chrome(service=s)
driver.maximize_window()
driver.get("https://www.walmart.com/ip/Clorox-Disinfecting-Wipes-225-Count-Value-Pack-Crisp-Lemon-and-Fresh-Scent-3-Pack-75-Count-Each/14898365")
driver.implicitly_wait(10)
clickobj=driver.find_element(By.XPATH,"//*[@id='customer-reviews-header']/div[2]/div/div[3]/a[2]/span")
clickobj.click()
lick=driver.find_element(By.XPATH,"/html/body/div[1]/div/div/div/div[1]/div/div[5]/div/div[2]/div/div[2]/div/div[2]/select/option[3]")
lick.click()
myreviews=[]
mydates=[]
mynames=[]
mytitles=[]
headers = {'User-Agent': 'Mozilla/5.0'}
for page in range(1,24):
   driver.get("https://www.walmart.com/reviews/product/14898365?sort=submission-desc"+"&page="+str(page))
   review = driver.find_elements_by_class_name("review-body")
   dates = driver.find_elements(By.XPATH, "//span[contains(@class,'submissionTime')]")
   names = driver.find_elements(By.XPATH, "//span[contains(@class,'userNickname')]")
   title = driver.find_elements(By.XPATH, "//h3[(@class='review-title font-bold')]")

   for reviews in review:
      myreviews.append(reviews.text)
   for data in dates:
      mydates.append(data.text)
   for name in names:
      mynames.append(name.text)
   for titles in title:
      mytitles.append(titles.text)
finalist=zip(mydates,mynames,myreviews,mytitles)
print("part 1")
wb=Workbook()
wb['Sheet'].title='Walmart'
sh1=wb.active
sh1.append(['Dates','Names','Reviews','Title','Ratings'])
for x in list(finalist):
    sh1.append(x)

wb.save("output.xlsx")
print("Part 2")
