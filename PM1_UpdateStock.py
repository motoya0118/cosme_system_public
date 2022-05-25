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
	with open(name,"a",newline="",encoding = 'shift-jis') as f:
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

def au_stock_update(au_api_key,ItemID,quantity):
	
	au_body = """
		<stockUpdateItem>
		<itemCode>{product_code}</itemCode>
		<stockSegment>1</stockSegment>
		<stockCount>{quantity}</stockCount>
		</stockUpdateItem>""".format(product_code=ItemID,quantity=quantity)
	
	requestURL = "https://api.manager.wowma.jp/wmshopapi/updateStock"
	xmlPostBody = """<?xml version="1.0" encoding="UTF-8"?>
	<request>
	<shopId>#ストアID(数字)を記入</shopId>
		{au_body}
	</request>""".format(au_body=au_body)
	#ここにXML形式でデータを記入
	headers = {"Content-type": "application/xml;charset=UTF-8","Authorization": "Bearer " + au_api_key}
	
	# POSTリクエスト送信
	bytesXMLPostBody = xmlPostBody.encode("UTF-8")
	req = requests.post(url=requestURL,data=bytesXMLPostBody,headers=headers)
	soup = BeautifulSoup(req.text, "xml")
	
	updateResults = soup.find_all("updateResult")
	
	for updateResult in updateResults:
		try:
			updateResult.find("error").text
			print(updateResult.find("itemCode").text,": 在庫更新に失敗しました")
		except:pass

def Qoo10UpdateStock(Qoo10_SAK,product_code,quantity):
	requestURL = "https://api.qoo10.jp/GMKT.INC.Front.QAPIService/ebayjapan.qapi"
	SellerCode = "co-1"
	Qty = "1"
	params = {
	"v":"1.0",
	"returnType":"xml",
	"method":"ItemsOrder.SetGoodsPriceQty",
	"key":Qoo10_SAK,
	"SellerCode":product_code,
	"Qty":quantity,
	}
	
	req = requests.get(url=requestURL,params=params)


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
			time.sleep(10)
	
	return driver	
def Yamato_stock_download():
	iDir = os.getcwd() + "/入力/YamatoFF_在庫"
	print(iDir)
	options = Options()
	options = webdriver.ChromeOptions()
	options.add_experimental_option("prefs", {"download.default_directory": iDir })
	options.add_argument('--headless')
	options.add_argument("--no-sandbox")
	options.add_argument("--remote-debugging-port=9222")
	driver = webdriver.Chrome(options=options)
	driver = Yamato_login(driver)
	
	url = "https://kuroneko-ylc.com/ffportal/Ebina/inventories"
	driver.get(url)
	time.sleep(3)
	flag = True
	while flag:
		try:
			driver.find_element_by_xpath("//*[@id='__layout']/div/section/div/section/div[2]/main/section/div/div[1]/div[2]/div/button").click()
			time.sleep(10)
			flag = False
		except:
			print("FalseClick")
	
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

def create_message_html(sender, to, subject, message_text,store_name):
	message = MIMEText(message_text,"html")
	message['to'] = to
	message['from'] = email.utils.formataddr((store_name, sender))
	message['subject'] = subject
	message.add_header('Content-Type','text/html')
	encode_message = base64.urlsafe_b64encode(message.as_bytes())
	return {'raw': encode_message.decode()}

def refresh_token(refresh_token):
	
	app_id = os.getenv("app_id")
	secret = os.getenv("secret")
	
	headers = {"Accept":"*/*","Content-Type":"application/x-www-form-urlencoded; charset=utf-8"}
	data = "grant_type=refresh_token&client_id={app_id}&client_secret={secret}&refresh_token={refresh_token}".format(app_id=app_id,secret=secret,refresh_token=refresh_token)
	
	url = "https://auth.login.yahoo.co.jp/yconnect/v2/token"

	req = requests.post(url=url,headers=headers,data=data)
	token_directry = json.loads(req.text)
	return token_directry["access_token"]

def Update_Stock(book_key,sh_list,refresh_token_base,au_api_key):	
	
	#ヤマトproduct_ID ➡ ストアproduct_IDに変換して在庫更新しないといけない	
	Yamato_stock_download()
	work_book = gspread_me(book_key)
	Listing_sheet = work_book.worksheet(sh_list[1])
	Listing_all = Listing_sheet.get_all_values()
	Listing_Maxrow = len(Listing_all) + 1
	key = "A1:T" + str(Listing_Maxrow)
	Listing_range = Listing_sheet.range(key)
	directry = {}
	directry_Co_P = {}
	#directory内容以前の処理と逆
	
	for L_row in Listing_all:
		if L_row[18] == "TRUE":	
			directry[L_row[14]] =  L_row[5]
			directry_Co_P[L_row[5]] = L_row[0]
	
	input_csv_path = glob_csv_path("./入力/YamatoFF_在庫")
	input_csvs = []
	input_csvs = read_csv(input_csv_path)
	
	exception_csv_path = "./入力/YamatoFF_在庫/例外リスト/例外リスト.csv"
	exception_csvs = []
	exception_csvs = read_csv(exception_csv_path)
	
	now = datetime.datetime.now()
	
	yahoo_filename = './出力_csv/' + "yahoo_在庫更新" + now.strftime('%Y%m%d_%H%M%S') + '.csv'
	yahoo_header = [["code","sub-code","quantity","mode"]]
		
	write_csv(yahoo_filename,yahoo_header)
	
	au_body = ""
	size_directory = {}
	P_To_Set_dic = {}
	Set_To_quontity_dic = {}
	ct = 0
	for i, row in enumerate(input_csvs):
		if i == 0:
			continue
		
		flag = False
		for exception_code in exception_csvs[0]:
			if row[0] == exception_code:
				flag = True
		
		if flag:
			continue
		
		product_code = directry.pop(row[0])
		quantity = row[14]
		
		if product_code == "co-4":
			quantity = int(quantity) -1
		
		P_To_Set_dic[directry_Co_P[product_code]] = quantity
		size_directory[row[0]] = row[9]
		
		yahoo_csv = [[product_code,"",quantity,""]]
		write_csv(yahoo_filename,yahoo_csv)
		au_body += """
		<stockUpdateItem>
		<itemCode>{product_code}</itemCode>
		<stockSegment>1</stockSegment>
		<stockCount>{quantity}</stockCount>
		</stockUpdateItem>""".format(product_code=product_code,quantity=quantity)
		ct += 1
		if ct == 50:
			requestURL = "https://api.manager.wowma.jp/wmshopapi/updateStock"
			xmlPostBody = """<?xml version="1.0" encoding="UTF-8"?>
			<request>
			<shopId>64197720</shopId>
				{au_body}
			</request>""".format(au_body=au_body)
			#ここにXML形式でデータを記入
			headers = {"Content-type": "application/xml;charset=UTF-8","Authorization": "Bearer " + au_api_key}
	
			# POSTリクエスト送信
			bytesXMLPostBody = xmlPostBody.encode("UTF-8")
			req = requests.post(url=requestURL,data=bytesXMLPostBody,headers=headers)
			soup = BeautifulSoup(req.text, "xml")
	
			updateResults = soup.find_all("updateResult")
			ct = 0
			au_body = ""
			for updateResult in updateResults:
				try:
					updateResult.find("error").text
					print(updateResult.find("itemCode").text,": 在庫更新に失敗しました")
				except:pass
	
		
		Qoo10UpdateStock(Qoo10_SAK,product_code,quantity)
	
	#セットリスト処理
	SET_Sheet = work_book.worksheet(sh_list[7])
	SET_all = SET_Sheet.get_all_values()
	
	auBodys = []

	for n,row in enumerate(SET_all):
		if n == 0:
			continue
		
		if not row[18] == "TRUE":
			continue
		
		Set_To_quontity_dic[row[0]] = []
		for i in range(5):
			if len(row[1+i*2]) > 0:
				try:Set_To_quontity_dic[row[0]].append(int(int(P_To_Set_dic[row[1+i*2]])/int(row[2+i*2])))
				except:Set_To_quontity_dic[row[0]].append(0)
		
		product_code = row[0]
		quantity = min(Set_To_quontity_dic[row[0]])
		
		yahoo_csv = [[product_code,"",quantity,""]]
		write_csv(yahoo_filename,yahoo_csv)
		au_body += """
		<stockUpdateItem>
		<itemCode>{product_code}</itemCode>
		<stockSegment>1</stockSegment>
		<stockCount>{quantity}</stockCount>
		</stockUpdateItem>""".format(product_code=product_code,quantity=quantity)
		ct += 1
		if ct == 50:
			requestURL = "https://api.manager.wowma.jp/wmshopapi/updateStock"
			xmlPostBody = """<?xml version="1.0" encoding="UTF-8"?>
			<request>
			<shopId>64197720</shopId>
				{au_body}
			</request>""".format(au_body=au_body)
			#ここにXML形式でデータを記入
			headers = {"Content-type": "application/xml;charset=UTF-8","Authorization": "Bearer " + au_api_key}
	
			# POSTリクエスト送信
			bytesXMLPostBody = xmlPostBody.encode("UTF-8")
			req = requests.post(url=requestURL,data=bytesXMLPostBody,headers=headers)
			soup = BeautifulSoup(req.text, "xml")
	
			updateResults = soup.find_all("updateResult")
			ct = 0
			au_body = ""
			for updateResult in updateResults:
				try:
					updateResult.find("error").text
					print(updateResult.find("itemCode").text,": 在庫更新に失敗しました")
				except:pass
		
		Qoo10UpdateStock(Qoo10_SAK,product_code,quantity)		
	
	for out_of_stock_ID in directry.values():
		product_code = out_of_stock_ID
		
		flag = False
		for exception_code in exception_csvs[0]:
			if product_code == exception_code:
				flag = True
		
		if flag:
			continue
		
		quantity = "0"
		
		yahoo_csv = [[product_code,"",quantity,""]]
		write_csv(yahoo_filename,yahoo_csv)
		au_body += """
		<stockUpdateItem>
		<itemCode>{product_code}</itemCode>
		<stockSegment>1</stockSegment>
		<stockCount>{quantity}</stockCount>
		</stockUpdateItem>""".format(product_code=product_code,quantity=quantity)
		Qoo10UpdateStock(Qoo10_SAK,product_code,quantity)
	
	#yahoo_api処理
	access_token = refresh_token(refresh_token_base)
	CSV_MIMETYPE = 'application/octet-stream'
	with open(yahoo_filename,mode='rb') as f:
		fileDataBinary = f.read()
    
	files = {'file':(yahoo_filename,fileDataBinary,'text/csv')}
	url = "https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/uploadStockFile?seller_id=chan-mu"
	headers = {"Authorization":"Bearer " + access_token}
	
	response = requests.post(url,headers=headers,files=files)
	print(response.text)
	
	#au_api処理
	requestURL = "https://api.manager.wowma.jp/wmshopapi/updateStock"
	xmlPostBody = """<?xml version="1.0" encoding="UTF-8"?>
	<request>
	<shopId>#ストアID(数字)を記入</shopId>
		{au_body}
	</request>""".format(au_body=au_body)
	#ここにXML形式でデータを記入
	headers = {"Content-type": "application/xml;charset=UTF-8","Authorization": "Bearer " + au_api_key}
	
	# POSTリクエスト送信
	bytesXMLPostBody = xmlPostBody.encode("UTF-8")
	req = requests.post(url=requestURL,data=bytesXMLPostBody,headers=headers)
	soup = BeautifulSoup(req.text, "xml")
	
	updateResults = soup.find_all("updateResult")
	for updateResult in updateResults:
		try:
			updateResult.find("error").text
			print(updateResult.find("itemCode").text,": 在庫更新に失敗しました")
		except:pass
	
	for z in range(Listing_Maxrow):
		if z == 0:
			continue
		
		try:
			size = size_directory[Listing_range[z * 20+14].value]
			Listing_range[z * 20 + 19].value = size
		except:pass
	
	Listing_sheet.update_cells(Listing_range)
	shutil.move(input_csv_path,"./入力/YamatoFF_在庫/処理済")
	print("Update_Stock_Finish")

au_api_key = os.getenv("au_api_key")
refresh_token_base = os.getenv("refresh_token_base")
Qoo10_SAK = os.getenv("Qoo10_SAK")
book_key = os.getenv("CosmeProduct")
sh_list = ["仕入れ商品リスト","出品商品リスト","仕入","引当","在庫","受注","新規注文","セット商品リスト"]

#Moniterの一次データを空にする
os.remove("./Moniter/Memory_Order.csv")
os.remove("./Moniter/Do_Zero.csv")
Memory_Order_path = pathlib.Path("./Moniter/Memory_Order.csv")
Memory_Order_path.touch()
DO_ZERO_path = pathlib.Path("./Moniter/Do_Zero.csv")
DO_ZERO_path.touch()

Update_Stock(book_key,sh_list,refresh_token_base,au_api_key)