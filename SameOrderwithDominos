from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import getpass
import os
#Sign in
eml = input('Email: ')
passwd = getpass.getpass('Password: ')

#optionally, if using the same login everytime, 
#use the variables below and comment out the top variables.

#eml = os.environ.get('DOMUSER')
#passwd = os.environ.get('DOMPASS')
#print (passwd)

 driver = webdriver.Firefox()
 driver.implicitly_wait(5) seconds
 url = 'https://www.dominos.com/en/pages/customer/!/customer/login/'
 driver.get(url)



 driver.find_element_by_id('Email').send_keys(eml)
 driver.find_element_by_id('Password').send_keys(passwd)
 driver.find_element_by_css_selector('button.btn.btn--small.js-loginSubmit').submit()					

 driver.switch_to.window(driver.window_handles[0])
 element = WebDriverWait(driver,15).until(EC.presence_of_element_located((By.CSS_SELECTOR,'div.js-listOfProducts')))Page after login

 items = driver.find_elements_by_css_selector('div.grid--duo-list__item__content')

 print('=' * 100)
 #Listing East Order Here for quick verification before ordering
 print('Here is the current easy order set. IF YOU WANT TO CHANGE THE ORDER, YOU NEED TO MANUALLY SIGN IN ONLINE TO MAKE ADJUSTMENTS. I always carryout and I pay at store. ')
 print('=' * 100)
 for item in items:
 	print('-' * 50)
 	print(item.text)

 print('*'*100)
 response =input('TO CONFIRM THIS ORDER AND TO SEE ORDER SUMMARY AND STATUS OF ORDER , ENTER \'y\' OR PRESS ENTER TO EXIT: ')
 print('*'*100)

 if response == 'y':
 	print ('ORDERING PIZZA. PLEASE WAIT ...')
 	driver.find_element_by_class_name('js-orderThis.btn.order-reorder-btn.fr.js-easyOrder').click()clicked reorder easy order
 	driver.switch_to.window(driver.window_handles[0])
 	if driver.find_element_by_class_name('card.card--overlay.js-cardOverlay'):
 		driver.find_element_by_css_selector('a.btn.btn--small.js-continueServiceMethodSwappedOverlay').click()
 	driver.switch_to.window(driver.window_handles[0])
 	driver.implicitly_wait(10) seconds
 	driver.find_element_by_class_name('btn.btn--large.btn--block.btn--continue-checkout.submitButton.qa-OrCheck.js-continueCheckout.c-order-continueCheckout').click()clicked Continue Checkout
	
 	nothanks = ['js-nothanks.btn.btn--outline','js-nothanks.btn.btn--secondary','js-nothanks.upsell--note']
 	driver.implicitly_wait(10) seconds

 	if driver.find_element_by_id('genericOverlay') or driver.find_element_by_id('overlayUIBlock2'):
 		for x in range(len(nothanks)):
 			try:
 				driver.find_element_by_class_name(nothanks[x]).click()
 			except Exception:
 				continue

 	driver.implicitly_wait(10) seconds
 	driver.find_element_by_class_name('c-order-payment-cash.js-paymentType').click()
	



 else:
 	driver.find_element_by_class_name('site-nav__profile__not-user.js-notUser.c-header-sign-out').click()
 	driver.quit()
 	print ('exited')
