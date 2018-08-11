from selenium import webdriver
import win32com.client
from selenium.webdriver.common.action_chains import ActionChains
import time
from selenium.webdriver.common.keys import Keys


browser = webdriver.Ie()
actions = ActionChains(browser)


browser.get("https://www.ctracker.ford.com/ConcernTrackerUiWeb/home.faces?dswid=2814")
Login = browser.find_element_by_id("ADloginUserIdInput")
Password = browser.find_element_by_id("ADloginPasswordInput")
Button = browser.find_element_by_id("ADloginWSLSubmitButton")
CDSID = raw_input("CDSID:")
Pass = raw_input("Password:")
browser.execute_script("arguments[0].value = arguments[2]; arguments[1].value = arguments[3];", Login, Password, CDSID, Pass)
Button.click()
Button.click()

W = 0

while W == 0:
	try:
		link = browser.find_element_by_xpath("//a[contains(@href,'searchSummary.faces')]")
		#Buttons = browser.find_elements_by_class_name("ui-menuitem ui-widget ui-corner-all")
		#for button in Buttons:
		#	if button.get_attribute("text") = "S
		W = 1
	except:
		time.sleep(2)
	
link.click()

time.sleep(2)

ModelCode = browser.find_element_by_id("SearchSummaryForm:accPanel:ptvl")
browser.execute_script("arguments[0].value = 'CVE1';", ModelCode)

ModelYears = browser.find_element_by_id("SearchSummaryForm:accPanel:modelYears")

for option in ModelYears.find_elements_by_tag_name("option"):
	if not option.get_attribute("selected"):
		option.click()

ReportClasses = browser.find_element_by_id("SearchSummaryForm:accPanel:rptclasses")

for option in ReportClasses.find_elements_by_tag_name("option"):
	if not option.get_attribute("selected"):
		option.click()
		

time.sleep(2)

SubmitSearch = browser.find_element_by_id("SearchSummaryForm:accPanel:submitsearch")
SubmitSearch.click()

time.sleep(2)
pageHTML = browser.execute_script("return document.body.innerHTML")
F1 = pageHTML.find("Please wait while your request is being processed")
F2 = F1
while F2 == F1:
	time.sleep(2)
	try: 
		pageHTML = browser.execute_script("return document.body.innerHTML")
		F2 = pageHTML.find("Please wait while your request is being processed")
		time.sleep(2)
	except:
		time.sleep(1)
	
Excel = browser.find_element_by_class_name("excelLink")
browser.execute_script("return arguments[0].scrollIntoView();", Excel)
actions.move_to_element(Excel).click().perform()

autoit = win32com.client.Dispatch("AutoItX3.Control")

time.sleep(2)
autoit.ControlFocus("Search Results - Internet Explorer", "", "DirectUIHWND1")
time.sleep(5)
autoit.ControlClick("Search Results - Internet Explorer", "", "DirectUIHWND1")
time.sleep(3)
autoit.ControlSend("Search Results - Internet Explorer", "", "DirectUIHWND1", "{F6}")
time.sleep(3)
autoit.ControlSend("Search Results - Internet Explorer", "", "DirectUIHWND1", "{TAB}")
time.sleep(3)
autoit.ControlSend("Search Results - Internet Explorer", "", "DirectUIHWND1", "{DOWN}")
time.sleep(3)
autoit.ControlSend("Search Results - Internet Explorer", "", "DirectUIHWND1", "{DOWN}")
time.sleep(3)
autoit.ControlSend("Search Results - Internet Explorer", "", "DirectUIHWND1", "{ENTER}")
print "End of clicks"
time.sleep(2)
	
autoit.ControlSetText("Save As", "", "Edit1", "")
autoit.ControlSend("Save As", "", "Edit1", "C:\Users\QR Admin\Downloads\ConcernTracker.xls")
autoit.ControlClick("Save As", "", "Button1")

time.sleep(2)

try:
	autoit.ControlClick("Confirm Save As", "", "Button1")
	print "Concern Tracker download complete"
except:
	print "Concern Tracker download complete"
time.sleep(2)
	
browser.quit()
#Turn into a function can import with a login, search, optional list, file path/name as inputs

