import os
import selenium
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import xlrd

chrome_driver = os.path.abspath(r"C:\Program Files (x86)\Google\Chrome\Application\chromedriver.exe");
os.environ["webdriver.chrome.driver"] = chrome_driver;
driver = webdriver.Chrome(chrome_driver);

#options = webdriver.ChromeOptions()
#options.add_argument('--ignore-certificate-errors')
#driver = webdriver.Chrome(executable_path=chrome_driver,chrome_options=options)

def Pay(cardno,amount):
	result=0;
	try:
		driver.get("登陆成功后的地址");
		time.sleep(1);
		elem = driver.find_element_by_id("ctl00_body_amount");
		elem.send_keys(amount);
		time.sleep(5);
		iswhile=True;
		while (iswhile):
			try:
				elem = driver.find_element_by_css_selector(".lk.FL.fw.ml20.cur.bkc");
				iswhile = False;
				elem.click();
			except Exception, e:
				print("waitting");
		try:
			print("选卡");
			elem = driver.find_element_by_xpath("//div[contains(@class,'pageBankimg')]/p[contains(text()," + cardno +")]");
			
			elem.click();
			elem = driver.find_element_by_id("ctl00_body_txtPaypwd");
			elem.send_keys("输入验证信息");
			elem.send_keys(Keys.RETURN);
			time.sleep(1);
			for handle in driver.window_handles:
				driver.switch_to_window(handle)
			elem = driver.find_element_by_class_name("rst_body");
			if "银行卡可用余额" in elem.text:
				print ("支付结果：银行卡(" + str(cardno) +")可用余额(活期存款)不足("+ str(amount) +")" );
			else:
				print(elem.text)
			result=1;
		except Exception:
			print ("支付结果：银行卡(" + str(cardno) +")不在本账户下" );
	except BaseException as ex:
		print("异常结束:" + ex)
		result=0;
	return result;

def PayDeal(table):
	nrows = table.nrows # 获取表的行数
	print("开始");
	for i in range(1,2): # 循环逐行打印
		print("第"+str(i));
		Pay("7206",12345); 
		#driver.implicitly_wait(30)

#调用函数
driver.get("登陆网址");
elem = driver.find_element_by_id("tbname");
elem.send_keys("用户名");
elem = driver.find_element_by_id("tbpwd");
elem.send_keys("密码");
elem.send_keys(Keys.RETURN);
driver.implicitly_wait(10);
driver.get("登陆成功后的地址");

data = xlrd.open_workbook(r'D:\test.xls') # 打开xls文件
table = data.sheets()[0] # 打开第一张表
PayDeal(table);
#driver.close();
#driver.quit();
#driver = None;
