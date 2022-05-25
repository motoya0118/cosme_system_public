from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
import chromedriver_binary 
from bs4 import BeautifulSoup
import time, codecs, csv, sys,re,glob,os,pickle,difflib,os.path,base64,email.utils,gspread,datetime,requests,pathlib,shutil
from oauth2client.service_account import ServiceAccountCredentials
from dateutil.relativedelta import relativedelta
from tkinter import messagebox
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from email.mime.text import MIMEText
from apiclient import errors
import urllib.request
import json
import pprint
import schedule
from dotenv import load_dotenv
load_dotenv()


def glob_csv_path(path):

	files = glob.glob(os.path.join(path,"*.csv"))
	for n,file in enumerate(files):
		path = file
	if not n == 0:
		print("csvが複数存在します。処理を中止します。")
		sys.exit()
	return path

def Crome(driver,url):
	driver.get(url)
	time.sleep(2)

def bs4_read(selects_css):
	html_source_2 = driver.page_source
	soup = BeautifulSoup(html_source_2, 'html.parser')
	lis = []
	lis = soup.select(selects_css)
	return lis

def bs4_read_href(html,select_css):
	soup = html
	href = soup.select_one(select_css).attrs["href"]
	return href

def bs4_read_string(html,select_css):
	soup = html
	seller_id_find = soup.select_one(select_css).string
	return seller_id_find

def read_csv(path):
	result = []
	with open(path,errors='ignore') as f:
		reader = csv.reader(f)
		for row in reader:
			result.append(row)
	return result

def write_csv(name,list):
	with open(name,"a",newline="") as f:
		writer = csv.writer(f)
		for a in list:
			writer.writerow(a)

def write_Shift_jis_csv(name,list):
	with open(name,"a",newline="",encoding = 'cp932') as f:
		writer = csv.writer(f)
		for a in list:
			writer.writerow(a)
	


def format_date(a):
	a = a.split("/")
	a = datetime.date(int(a[0]),int(a[1]),int(a[2]))
	return a.strftime("%Y%m%d")

def au_settleStatus(settleStatus):
	True_list = ["AD","TD","CD","TS","NR"]
	check = False
	for True_character in True_list:
		if True_character == settleStatus:
			check = True
	
	return check

def return_size(size,hope):
	if size == "SS1":
		if hope == "":
			return "0501"
	
	return "0101"

def right_show(source,target_character):
	target = target_character				
	treat_type = source.find(target)
	right_show = source[treat_type+1:]
	return right_show


def gspread_me(book_key):
	scope = ['https://spreadsheets.google.com/feeds',
			'https://www.googleapis.com/auth/drive']

	credentials = ServiceAccountCredentials.from_json_keyfile_name('python-262511-503e8108476e.json', scope)
	gc = gspread.authorize(credentials)

	work_book = gc.open_by_key(book_key)
	return work_book

def Yamato_login(driver):
	
	url = "https://kuroneko-ylc.com/ffportal/Ebina/login"	
	#ログイン
	flag = True
	ct = 0
	while flag:
		try:
			driver.get(url)
			time.sleep(10)
			YamatoId = os.getenv("YamatoId")
			YamatoPass = os.getenv("YamatoPass")
			driver.find_element_by_xpath("//*[@id='__layout']/section/span/div/div/section/span[1]/div/div/input").send_keys(YamatoId)
			driver.find_element_by_xpath("//*[@id='__layout']/section/span/div/div/section/span[2]/div/div/input").send_keys(YamatoPass)
			driver.find_element_by_xpath("//*[@id='__layout']/section/span/div/div/section/button").click()
			time.sleep(10)
			flag = False
			if ct == 5:
				sys.exit("5回ログイン失敗したため終了")
		except:
			print("ログイン失敗")
			ct += 1
			flag = True
			time.sleep(10)
	
	return driver

def Yamato_product_uoload(upload_file):
	options = Options()
	options.add_argument('--headless')
	options.add_argument("--no-sandbox")
	options.add_argument("--remote-debugging-port=9222")
	driver = webdriver.Chrome(options=options)
	driver = Yamato_login(driver)
	#商品情報アップロード
	url = "https://kuroneko-ylc.com/ffportal/Ebina/products/bulk-register"
	driver.get(url)
	time.sleep(10)
	flag = True
	while flag:
		try:
			driver.find_element_by_xpath("//*[@id='__layout']/div/section/div/section/div[2]/main/section/div/div[1]/label/input").send_keys(upload_file)
			driver.find_element_by_xpath('//*[@id="__layout"]/div/section/div/section/div[2]/main/section/div/div[3]/p/button').click()
			time.sleep(10)
			driver.find_element_by_xpath('/html/body/div[2]/div[2]/footer/button[2]').click()
			time.sleep(5)
			print("END")
			flag = False
		except:
			print("uploadFalse")
			driver.get(url)
			time.sleep(5)

def yahoo_stock_update(access_token,ItemID,quantity):
	#商品IDと在庫数を代入してヤフーショッピングの在庫を更新する
	#在庫更新URL
	quontity_url = "https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/setStock"
	headers = {"Authorization":"Bearer " + access_token}
	data = {"seller_id":"chan-mu","item_code":ItemID,"quantity":quantity}
	req = requests.post(url=quontity_url,headers=headers,data=data)
	soup = BeautifulSoup(req.text, "xml")
	print(soup)
			
	if not req.status_code == 200:
		print("在庫更新失敗",soup)


def chatwork_message(chatwork_api_key,message):
	room_id = #ルームIDを指定
	url = "https://api.chatwork.com/v2/rooms/{room_id}/messages".format(room_id=room_id)
	headers = {"X-ChatWorkToken":chatwork_api_key}	
	params = {"body":message,"self_unread":1}

		
	req = requests.post(url=url,headers=headers,params=params)
	if not req.status_code == 200:
		print(req.text)
		sys.exit("メッセージ投稿失敗")


def Change_SetID_to_YamatoID(work_book,sh_list,Directory_Listing_Code,Directory_Listing_size,Directory_Listing_Title,SetID,quantity):
	SET_Sheet = work_book.worksheet(sh_list[7])
	SET_all = SET_Sheet.get_all_values()
	
	Set_To_P_list = []
	Set_To_Yamato_list = []
	for row in SET_all:
		if row[0] == SetID:
			for i in range(5):
				if len(row[1 + i*2]) > 0:
					Set_To_P_list.append([row[1 + i*2],row[2 + i*2]])
		
			for i in Set_To_P_list:
				Yamato_ID = Directory_Listing_Code[i[0]]
				size = Directory_Listing_size[i[0]]
				Title = Directory_Listing_Title[i[0]]
				Set_To_Yamato_list.append([Yamato_ID,int(i[1])*int(quantity),size,Title])
			
			return Set_To_Yamato_list
			
	sys.exit("SET商品が商品リストに存在しません。処理を中止します。")

def Change_SetID_to_ListingID(work_book,sh_list,Directory_Listing_Code,Directory_Listing_Price,SetID,quantity):
	SET_Sheet = work_book.worksheet(sh_list[7])
	SET_all = SET_Sheet.get_all_values()
	
	Set_To_P_list = []
	Set_To_Yamato_list = []
	for row in SET_all:
		if row[0] == SetID:
			for i in range(5):
				if len(row[1 + i*2]) > 0:
					Set_To_P_list.append([row[1 + i*2],row[2 + i*2]])
			
			#product_IDを入れるとListing_IDが返ってくる
			for i in Set_To_P_list:
				Price = Directory_Listing_Price[i[0]]
				Set_To_Yamato_list.append([i[0],int(i[1])*int(quantity),Price])
			
			return Set_To_Yamato_list
	
	sys.exit("SET商品が商品リストに存在しません。処理を中止します。")

def refresh_token(refresh_token):
	
	app_id = os.getenv("app_id")
	secret = os.getenv("secret")
	
	headers = {"Accept":"*/*","Content-Type":"application/x-www-form-urlencoded; charset=utf-8"}
	data = "grant_type=refresh_token&client_id={app_id}&client_secret={secret}&refresh_token={refresh_token}".format(app_id=app_id,secret=secret,refresh_token=refresh_token)
	
	url = "https://auth.login.yahoo.co.jp/yconnect/v2/token"

	req = requests.post(url=url,headers=headers,data=data)
	token_directry = json.loads(req.text)
	return token_directry["access_token"]

def Normal_Listing(book_key,sh_list,refresh_token_base,chatwork_api_key):
	
	##出品商品リスト側
	work_book = gspread_me(book_key)
	work_sheet = work_book.worksheet(sh_list[1])
	get_all = work_sheet.get_all_values()
	count = len(get_all)
	flag_cols = "S1:S" + str(count)
	flag_col = work_sheet.range(flag_cols)
	
	now = datetime.datetime.now()
	
	#ヤフーショッピング
	csv_path = "./入力/csv雛形/data_input202201101638.csv"
	template = []
	template = read_csv(csv_path)
	
	#ヤマト
	yamato_csv_path = "./入力/csv雛形/ヤマトFF.csv"
	yamato_template = []
	yamato_template = read_csv(yamato_csv_path)
	yamato_filename = './出力_csv/' + "ヤマト" + now.strftime('%Y%m%d_%H%M%S') + '.csv'
	
	filename = './出力_csv/' + "出品" + now.strftime('%Y%m%d_%H%M%S') + '.csv'
	header = [template[0]]
	write_Shift_jis_csv(filename,header)
	
	#出力ログ
	log_file_name = './ログ/' + "Listing_" + now.strftime('%Y%m%d_%H%M%S') + '.csv'
	log = []
	
	flag = False
	
	for i, infoes in enumerate(get_all):
		if i == 0:
			continue
		
		if len(infoes[18]) == 0:
			if len(infoes[14]) > 0 and len(infoes[15]) > 0 and len(infoes[16]) > 0 and len(infoes[17]) > 0:
				
				#ヤフーショッピング側csv出力処理
				product_code = infoes[5]
				product_name = infoes[6]
				catch_copy = infoes[7]
				JAN = infoes[9]
				price = infoes[10].replace(",","")
				explanation = infoes[11]
				product_categiry = infoes[12]
				yamato_ID = infoes[14]
				yamato_conform_no = infoes[15]
				yamato_type = infoes[16]
				yamato_into = infoes[17]
				reference_URL = infoes[13]
				brand_code = "38074"
				
				template_body = []
				template_body = template[1]
				
				template_body[1] = product_name
				template_body[2] = product_code
				template_body[5] = price
				template_body[8] = catch_copy
				template_body[11] = explanation
				template_body[27] = brand_code
				template_body[30] = JAN
				template_body[35] = product_categiry
				
				fake_list = [template_body]
				print(fake_list)
				write_Shift_jis_csv(filename,fake_list)
				
				flag_col[i].value = "TRUE"
				
				#yamato_csv出力処理
				yamato_body = []
				yamato_body = yamato_template[1]
				yamato_body[0] = yamato_ID
				yamato_body[1] = yamato_conform_no
				yamato_body[2] = product_name
				
				if yamato_type == "ピース":
					yamato_type = 1
				elif yamato_type == "ボール":
					yamato_type = 2
				else:
					yamato_type = 3
				
				yamato_body[11] = yamato_type
				yamato_body[10] = yamato_into
				yamato_body[18] = product_code
				
				yamato_fake = [yamato_body]
				write_csv(yamato_filename,yamato_fake)
				flag = True
				
				URL = "https://store.shopping.yahoo.co.jp/chan-mu/" + product_code
				log.append(product_code)
				message = """
				===出品完了===
				※画像未登録の場合は登録お願いします※
				商品コード：{product_code}
				商品名：{product_name}
				出品URL：{URL}
				///////////
				出品参考URL：{reference_URL}
				============
				""".format(product_code=product_code,product_name=product_name,URL=URL,reference_URL=reference_URL)
				chatwork_message(chatwork_api_key,message)
				
	if flag == True:
		#ヤマト側のcsvアップロード
		pathlib_filename = pathlib.Path(yamato_filename)
		upload_file = str(pathlib_filename.resolve())
		print(upload_file)
		Yamato_product_uoload(upload_file)
		
		#ヤフーショッピング側のcsvアップロード
		access_token = refresh_token(refresh_token_base)
		CSV_MIMETYPE = 'application/octet-stream'
		with open(filename,mode='rb') as f:
			fileDataBinary = f.read()
		
		files = {'file':(filename,fileDataBinary,'text/csv')}
		url = "https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/uploadItemFile?seller_id=chan-mu"
		
		headers = {"Authorization":"Bearer " + access_token}	
		data = {"type":"1"}
		req = requests.post(url,headers=headers,data=data,files=files)
		
		print(req.text)
		
		for Item_Code in log:
			yahoo_stock_update(access_token,Item_Code,0)
		
	write_csv(log_file_name,[log])
	work_sheet.update_cells(flag_col)
	print("Normal_Listing_Finish")
	
def SET_Listing(book_key,sh_list,refresh_token_base,chatwork_api_key):
	
	##出品商品リスト側
	work_book = gspread_me(book_key)
	work_sheet = work_book.worksheet(sh_list[7])
	get_all = work_sheet.get_all_values()
	count = len(get_all)
	flag_cols = "S1:S" + str(count)
	flag_col = work_sheet.range(flag_cols)
	
	now = datetime.datetime.now()
	
	#ヤフーショッピング
	csv_path = "./入力/csv雛形/data_input202201101638.csv"
	template = []
	template = read_csv(csv_path)
	
	filename = './出力_csv/' + "SET出品" + now.strftime('%Y%m%d_%H%M%S') + '.csv'
	header = [template[0]]
	write_Shift_jis_csv(filename,header)
	
	#出力ログ
	log_file_name = './ログ/' + "SET_Listing_" + now.strftime('%Y%m%d_%H%M%S') + '.csv'
	log = []
	
	flag = False
	
	for i, infoes in enumerate(get_all):
		if i == 0:
			continue
		
		if len(infoes[18]) == 0:
			if len(infoes[14]) > 0 and len(infoes[15]) > 0 and len(infoes[16]) > 0:
				
				#ヤフーショッピング側csv出力処理
				product_code = infoes[0]
				product_name = infoes[11]
				catch_copy = infoes[12]
				JAN = infoes[13]
				price = infoes[14].replace(",","")
				explanation = infoes[15]
				product_categiry = infoes[16]
				reference_URL = infoes[17]
				
				template_body = []
				template_body = template[1]
				
				template_body[1] = product_name
				template_body[2] = product_code
				template_body[5] = price
				template_body[8] = catch_copy
				template_body[11] = explanation
				template_body[30] = JAN
				template_body[35] = product_categiry
				
				fake_list = [template_body]
				print(fake_list)
				write_Shift_jis_csv(filename,fake_list)
				
				flag_col[i].value = "TRUE"
				
				URL = "https://store.shopping.yahoo.co.jp/chan-mu/" + product_code
				log.append(product_code)
				message = """
				===出品完了===
				※画像未登録の場合は登録お願いします※
				商品コード：{product_code}
				商品名：{product_name}
				出品URL：{URL}
				///////////
				出品参考URL：{reference_URL}
				============
				""".format(product_code=product_code,product_name=product_name,URL=URL,reference_URL=reference_URL)
				chatwork_message(chatwork_api_key,message)
				flag = True
				
	if flag == True:
		
		#ヤフーショッピング側のcsvアップロード
		access_token = refresh_token(refresh_token_base)
		CSV_MIMETYPE = 'application/octet-stream'
		with open(filename,mode='rb') as f:
			fileDataBinary = f.read()
		
		files = {'file':(filename,fileDataBinary,'text/csv')}
		url = "https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/uploadItemFile?seller_id=chan-mu"
		
		headers = {"Authorization":"Bearer " + access_token}	
		data = {"type":"1"}
		req = requests.post(url,headers=headers,data=data,files=files)
		
		print(req.text)
		
		for Item_Code in log:
			yahoo_stock_update(access_token,Item_Code,0)
		
	write_csv(log_file_name,[log])
	work_sheet.update_cells(flag_col)
	print("Listing_Finish")


def Master_Listing(book_key,sh_list,refresh_token_base,chatwork_api_key):
	
	Normal_Listing(book_key,sh_list,refresh_token_base,chatwork_api_key)
	SET_Listing(book_key,sh_list,refresh_token_base,chatwork_api_key)
	print("Master_Listing_Finish")	

au_api_key = os.getenv("au_api_key")
refresh_token_base = os.getenv("refresh_token_base")
chatwork_api_key = os.getenv("ChatWorkApiKey")
book_key = os.getenv("CosmeProduct")
sh_list = ["仕入れ商品リスト","出品商品リスト","仕入","引当","在庫","受注","新規注文","セット商品リスト"]

Master_Listing(book_key,sh_list,refresh_token_base,chatwork_api_key)