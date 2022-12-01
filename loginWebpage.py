from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
import pandas as pd
import time
import yaml

conf = yaml.safe_load(open('loginDetails.yml'))
id = conf['aks_user']['id']
ids = conf['aks_user']['dids']

myFbPassword = conf['aks_user']['password']
#URL = 'http://{host_name}/QuadraPro/QuadraPortal/Client/'
URL = 'login_page_url'
LOGIN_ROUTE = 'ClientPortalLogin.aspx'
HOME_ROUTE = 'Home.aspx'

df = pd.DataFrame(columns=['name','ApartmentNo' ,'Mobile', 'email', 'Guardian'])

driver = webdriver.Chrome()

def login(url,usernameId, username, passwordId, password, submit_buttonId):
   driver.get(url)
   driver.find_element("name",usernameId).send_keys(username)
   driver.find_element("name",passwordId).send_keys(password)
   driver.find_element("name",submit_buttonId).click()


def extract_to_df(df):
    driver.switch_to.frame(driver.find_elements(By.TAG_NAME, "iframe")[0])
    name_element = driver.find_elements("xpath",'/html/body/form/div[3]/table[2]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr[2]/td[2]')
    name =name_element[0].get_attribute('innerText')

    apno_element = driver.find_elements("xpath",'/html/body/form/div[3]/table[2]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr[5]/td[2]')
    apno=apno_element[0].get_attribute('innerText')

    mobno_element = driver.find_elements("xpath",'/html/body/form/div[3]/table[2]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr[15]/td[2]')
    mobno=mobno_element[0].get_attribute('innerText')

    email_element = driver.find_elements("xpath",'/html/body/form/div[3]/table[2]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr[15]/td[5]')
    email=email_element[0].get_attribute('innerText')

    guardian_element = driver.find_elements("xpath",' /html/body/form/div[3]/table[2]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr[3]/td[2]')
    guardian=guardian_element[0].get_attribute('innerText')

    df.loc[len(df.index)] = [name, apno, mobno,email, guardian]



loginIds = ids.split(',')
for i in range(len(loginIds)):
    loginId =loginIds[i].strip()
    try:
        login(URL+LOGIN_ROUTE, "txtUsername", loginId, "txtPassword", myFbPassword, "ibtnLogin")
        wait = WebDriverWait( driver, 200 )
        time.sleep(3)

        extract_to_df(df)
    except IndexError as e:
            print(e)
            continue

df.to_excel("output.xlsx")




driver.close()
