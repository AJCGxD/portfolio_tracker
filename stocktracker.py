from selenium import webdriver
from selenium.webdriver.common.by import By
import random,time
import xlwings as xw
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


wb = xw.Book(r"G:\New folder\programs\python\tonk.xlsx")
sheet = wb.sheets['Sheet1']

with open("user agents.txt", 'r') as file:
    user_agents = file.readlines()
user_agents = [agent.strip() for agent in user_agents]
sur = random.choice(user_agents)

options = webdriver.FirefoxOptions()
#options.add_argument("--headless")
options.set_preference("general.useragent.override", sur)
browser = webdriver.Firefox(options=options)
list1 = []
company = ""

kite_list = []

print("Kite stonks")
stockno = int(input("Number of stocks: "))
for i in range(stockno):
    company = input("Enter the symbol of the company: ")
    kite_list.append(company)

print("\nFrontpage Stonks")
frontpage_list = []
stocknof = int(input("Number of stocks: "))
for i in range(stocknof):
    company = input("Enter the symbol of the company: ")
    frontpage_list.append(company)

wait = WebDriverWait(browser,10)
pricel = []

def update_price(stock_list, start_row, price_column):
    for cid, comp in enumerate(stock_list):
        url = f"https://in.tradingview.com/chart/?symbol=NSE:{comp}"
        browser.get(url)

        price_element = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, ".price-qWcO4bp9")))
        price = price_element.text

        row = start_row + cid
        sheet.range(f'{price_column}{row}').value = price
        time.sleep(random.uniform(5,10))
try:
    while True:
        print("Updating KITE")
        update_price(kite_list,2,'I')

        print("Updating Frontpage")
        update_price(frontpage_list,20 ,'I')

        wb.save()

        sleep_time = random.uniform(300,400)
        print(f"Sleeping for {sleep_time/60: .2f} minutes")
        time.sleep(sleep_time)

except KeyboardInterrupt:
    print("program stopped")
    browser.quit()



