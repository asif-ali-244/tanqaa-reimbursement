from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from pdfquery import PDFQuery
import glob, re

def readfiles():
    pdfs = []
    for file in glob.glob("*.pdf"):
        pdfs.append(file)
    return pdfs

def get_date(date):
    date_split = date.split()
    month = date_split[0]
    day = re.sub('\D', '', date_split[1])
    year = date_split[2]
    return (day, month, year)

def monthToNum(shortMonth):
    return {
            'Jan': 1,
            'Feb': 2,
            'Mar': 3,
            'Apr': 4,
            'May': 5,
            'Jun': 6,
            'Jul': 7,
            'Aug': 8,
            'Sep': 9,
            'Oct': 10,
            'Nov': 11,
            'Dec': 12
    }[shortMonth]

def generate_data_rapido():
    print("Generating rapido data")
    files = readfiles()
    f = open("summary.txt", "w")
    f.write("ride_id date amount")
    f.write('\n')
    count = 0
    rows = []
    for file in files:
        count += 1
        pdf = PDFQuery(file)
        pdf.load()
        pdf.tree.write('pdfXML.txt', pretty_print = True)
        date = pdf.pq('LTTextLineHorizontal:overlaps_bbox("407.965, 699.363, 506.626, 708.363")').text()
        amount = pdf.pq('LTTextLineHorizontal:overlaps_bbox("266.707, 654.598, 339.28, 676.797")').text()
        ride_id = pdf.pq('LTTextLineHorizontal:overlaps_bbox("399.211, 745.904, 517.881, 757.904")').text()
        # remove zero amount
        if amount.split()[-1] != '0.00':
            amount = amount.split()[-1].split(".")[0]
            date = date.split(",")[0]
            row = [ride_id, date, amount]
            rows.append(row)
            print(row)
            f.write(f"{ride_id} {date} {amount}")
            f.write('\n')
    f.write(f'Total Count: {count}')
    f.close()
    return rows

def main():
    rows = generate_data_rapido()
    options = webdriver.ChromeOptions()
    options.add_experimental_option('detach', True)
    options.add_argument("window-size=1200x600")

    url = "https://www.tibs.in/tanqaa/LoginValidate.do"
    global driver
    driver = webdriver.Chrome(options=options)
    driver.get(url)

    try:
        username = driver.find_element(By.ID, 'userId')
        company = driver.find_element(By.ID, 'companyName')
        password = driver.find_element(By.ID, 'userPassword')
        captcha = driver.find_element(By.ID, 'captchaValue')
        username_input = input("Enter USERNAME and Press ENTER\n")
        password_input = input("Enter PASSWORD and Press ENTER\n")
        company_input = input("Enter COMPANY and Press ENTER\n")
        captcha_input = input("Enter CAPTCHA and Press ENTER\n")

        username.send_keys(username_input)
        password.send_keys(password_input)
        company.send_keys(company_input)
        captcha.send_keys(captcha_input)

        submit_button = driver.find_element(By.CSS_SELECTOR, "input[type='submit']")
        submit_button.click()

        driver.switch_to.frame(driver.find_element(By.ID,'iframe'))
        driver.switch_to.frame(driver.find_element(By.XPATH, "//frame[@id='topmenu']"))
        driver.find_element(By.CSS_SELECTOR, ".a5[href='ReimbursementRelated.do']").click()
        driver.implicitly_wait(5)
        driver.switch_to.default_content()
        driver.switch_to.frame(driver.find_element(By.ID,'iframe'))
        driver.switch_to.frame(driver.find_element(By.XPATH, "//frame[@id='leftmenu']"))
        driver.find_element(By.CSS_SELECTOR, ".a4[href='EmployeeNonpayrollReimbursementMenu.do']").click()

        driver.switch_to.default_content()

        driver.switch_to.frame(driver.find_element(By.ID,'iframe'))
        driver.switch_to.frame(driver.find_element(By.ID,'searchandresult'))
        driver.switch_to.frame(driver.find_element(By.XPATH, "//iframe[@id='contentIframe']"))
        driver.find_element(By.XPATH,"//button[@id='button1-button']").click()
        driver.implicitly_wait(5)
        driver.switch_to.default_content()

        driver.switch_to.frame(driver.find_element(By.ID,'iframe'))
        driver.switch_to.frame(driver.find_element(By.XPATH, "//frame[@id='heading']"))
        driver.find_element(By.XPATH, "//img[@title='Maximize']").click()
        driver.switch_to.default_content()
        driver.switch_to.frame(driver.find_element(By.ID,'iframe'))
        driver.switch_to.frame(driver.find_element(By.XPATH, "//frame[@id='content']"))

        count = 1
        for row in rows:
            print(row)
            ride_id = row[0]
            date = get_date(row[1])
            amount = row[2]
            # Click on add entry button
            driver.find_element(By.CSS_SELECTOR, "tbody tr td div div table thead tr th div span img[title='Add a blank row']").click()

            # Add Vendor detail
            driver.find_element(By.XPATH,f"//tbody/tr[{count}]/td[3]/div[1]").click()
            driver.find_element(By.XPATH,'//*[@id="yui-textboxceditor8-container"]/form/input').send_keys('Rapido')

            # Add Receipt no
            driver.find_element(By.XPATH,f"//tbody/tr[{count}]/td[4]/div[1]").click()
            driver.find_element(By.XPATH,'//*[@id="yui-textboxceditor9-container"]/form/input').send_keys(ride_id)

            # Add date
            driver.find_element(By.XPATH,f'//tbody/tr[{count}]/td[5]/div[1]').click()
            if date[1]!='Oct':
                driver.find_element(By.CLASS_NAME,"calnav").click()
                driver.implicitly_wait(5)
                Select(driver.find_element(By.CLASS_NAME,'yui-cal-nav-mc')).select_by_visible_text(date[1])
                driver.find_element(By.XPATH,"//button[@type='button'][normalize-space()='OK']").click()
            driver.implicitly_wait(5)
            month_num = monthToNum(date[1])
            driver.implicitly_wait(5)
            ele = driver.find_element(By.CLASS_NAME,f'm{month_num}')
            driver.implicitly_wait(5)
            tds = ele.find_elements(By.TAG_NAME, "a")
            for t in tds:
                if t.get_attribute('innerHTML') == str(date[0]):
                    t.click()
                    break


            # Add Amount
            driver.find_element(By.XPATH,f'//tbody/tr[{count}]/td[6]/div[1]').click()
            driver.find_element(By.XPATH,'//*[@id="yui-textboxceditor11-container"]/form/input').send_keys(amount)


            # Add reimburse type
            driver.find_element(By.XPATH,f'//tbody/tr[{count}]/td[2]/div[1]').click()
            select = Select(driver.find_element(By.XPATH,'//*[@id="yui-dropdownceditor7-container"]/select'))
            select.select_by_value("14")
            count+=1

        print("Done")

    except Exception as e:
        print(f"An error occurred: {str(e)}")

    # finally:
    #     driver.quit()


if __name__=="__main__":
    main()

