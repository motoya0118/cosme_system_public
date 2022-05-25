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
	
def Qoo10UpdateStock(Qoo10_SAK,ItemID,quantity):
	requestURL = "https://api.qoo10.jp/GMKT.INC.Front.QAPIService/ebayjapan.qapi"
	SellerCode = "co-1"
	Qty = "1"
	params = {
	"v":"1.0",
	"returnType":"xml",
	"method":"ItemsOrder.SetGoodsPriceQty",
	"key":Qoo10_SAK,
	"SellerCode":ItemID,
	"Qty":quantity,
	}
	
	req = requests.get(url=requestURL,params=params)
	print("Qoo10:",req.text)


def refresh_token(refresh_token):
	
	app_id = os.getenv("app_id")
	secret = os.getenv("secret")
	
	headers = {"Accept":"*/*","Content-Type":"application/x-www-form-urlencoded; charset=utf-8"}
	data = "grant_type=refresh_token&client_id={app_id}&client_secret={secret}&refresh_token={refresh_token}".format(app_id=app_id,secret=secret,refresh_token=refresh_token)
	
	url = "https://auth.login.yahoo.co.jp/yconnect/v2/token"

	req = requests.post(url=url,headers=headers,data=data)
	token_directry = json.loads(req.text)
	return token_directry["access_token"]

def Inventory_Monitoring(book_key,sh_list,refresh_token_base,au_api_key):
	#在庫0にした商品コードを入れる
	DO_ZERO_Path = "./Moniter/Do_Zero.csv"
	Do_ZERO_lists = read_csv(DO_ZERO_Path)
	Do_ZERO_list = []
	for row in Do_ZERO_lists:
		Do_ZERO_list.append(row[0])
	
	DO_ZERO_NEWS = []
	#受注管理表になく1度照会した注文番号を入れる
	Memory_Order_path = "./Moniter/Memory_Order.csv" 
	Memory_Order_lists = read_csv(Memory_Order_path)
	Memory_Order_list = []
	for row in Memory_Order_lists:
		Memory_Order_list.append(row[2])
	
	#スプレッドシート側
	work_book = gspread_me(book_key)
	Order_sheet = work_book.worksheet(sh_list[6])
	Order_all = Order_sheet.get_all_values()
	Order_ID_list = []
	for row in Order_all:
		Order_ID_list.append(row[0])

		#セットリスト
	SET_dic = {}
	SET_Sheet = work_book.worksheet(sh_list[7])
	SET_all = SET_Sheet.get_all_values()
	for row in SET_all:
		if row[18] == "TRUE":
			SET_dic[row[0]] = []
			for i in range(5):
				if len(row[1+i*2]) > 0:
					SET_dic[row[0]].append([row[1+i*2],int(row[2+i*2])])	
	
	#ヤマトcsv側
	Inventory_Files = glob.glob("./入力/YamatoFF_在庫/処理済/*.csv")
	Target_File = max(Inventory_Files,key=os.path.getctime)
	Yamato_csv = []
	Yamato_csv = read_csv(Target_File)
	
	#ヤマトコードを商品コードに変換
	Listing_sheet = work_book.worksheet(sh_list[1])
	Listing_all = Listing_sheet.get_all_values()
	Yamato_To_Product_dic = {}
	ProductSheetNo_To_ProductCode = {}
	Inventory_dic = {}
	for row in Listing_all:
		if row[18] == "TRUE":
			Yamato_To_Product_dic[row[14]] = row[5]
			ProductSheetNo_To_ProductCode[row[0]] = row[5]
		
	for i,row in enumerate(Yamato_csv):
		if i == 0:
			continue
		
		try:Inventory_dic[Yamato_To_Product_dic[row[0]]] = int(row[14])
		except:print("None:",row[0])
	
	#api側
		#下記のリストに差し引く商品コードと数量を入れていく
	New_Order_list = []
	
		#ヤフーショッピング側
	T_now = datetime.datetime.now()
	T_before = T_now + datetime.timedelta(days=-15)
	T_to = T_now.strftime('%Y%m%d%H%M%S')
	T_from = T_before.strftime('%Y%m%d%H%M%S')
	access_token = refresh_token(refresh_token_base)
		
	url = "https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/orderList"
	
	headers = {"Authorization":"Bearer " + access_token}	
		
	data = """
	<Req>
	<Search>
		<Result>1000</Result>
		<Condition>
		<OrderTimeFrom>{OrderTimeFrom}</OrderTimeFrom>
		<OrderTimeTo>{OrderTimeTo}</OrderTimeTo>
		</Condition>
		<Field>OrderId,ShipMethodName,OrderStatus,PayStatus</Field>
	</Search>
	<SellerId>chan-mu</SellerId>
	</Req>""".format(OrderTimeFrom=T_from,OrderTimeTo=T_to)
		
	req = requests.post(url=url,headers=headers,data=data,cert=(#crtファイルのパスを記入,#keyファイルのパスを記入))
	soup = BeautifulSoup(req.text, "xml")
	trees = []
	trees = soup.find_all("OrderInfo")
	Yahoo_New_Orders = []
	
	for row in trees:
		if row.find("ShipMethodName").text == #処理対象の配送方法名を記入:
			if row.find("OrderStatus").text == "2":
				if row.find("PayStatus").text == "1":
					if not row.find("OrderId").text in Order_ID_list:
						if not row.find("OrderId").text in Memory_Order_list:
							Yahoo_New_Orders.append(row.find("OrderId").text)
	
	for Yahoo_Order_Id in Yahoo_New_Orders:
		data = """
				<Req>
					<Target>
						<OrderId>{OrderId}</OrderId>
						<Field>OrderId,ShipMethodName,OrderStatus,PayMethodName,PayStatus,Discount,BillMailAddress,BillFirstName,BillLastName,ItemId,UnitPrice,SubCode,OrderTime,Quantity,UnitPrice,TotalPrice,ShipFirstName,ShipLastName,ShipMethod,ShipZipCode,ShipPrefecture,ShipCity,ShipAddress1,ShipAddress2,ShipPhoneNumber,ShipRequestDate,ShipRequestTime,Title,Discount,UsePoint,TotalPrice,TotalCouponDiscount,SettleAmount,ShipCharge,TotalMallCouponDiscount</Field>
					</Target>
					<SellerId>chan-mu</SellerId>
				</Req>""".format(OrderId = Yahoo_Order_Id)
				
		url = "https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/orderInfo"
		headers = {"Authorization":"Bearer " + access_token}
		req = requests.post(url=url,headers=headers,data=data,cert=(#crtファイルのパスを記入,#keyファイルのパスを記入))
		soup = BeautifulSoup(req.text, "xml")
		
		if not req.status_code == 200:
			sys.exit("リフレッシュトークンを更新して下さい")
		
		Items = soup.find_all("Item")
		for Item in Items:
			product_code = Item.find("ItemId").text
			Quantity = int(Item.find("Quantity").text)
			
			if product_code[:3] == "set" or product_code[:3] == "SET":
				for set in SET_dic[product_code]:
					New_Order_list.append([ProductSheetNo_To_ProductCode[set[0]],set[1]*Quantity,Yahoo_Order_Id])
			else:
				New_Order_list.append([product_code,Quantity,Yahoo_Order_Id])
	
		#au pay マーケット側
	au_url = "https://api.manager.wowma.jp/wmshopapi/"
	T_now = datetime.datetime.now()
	T_before = T_now + datetime.timedelta(days=-30)
	
	endDate = T_now.strftime('%Y%m%d')
	startDate = T_before.strftime('%Y%m%d')
	requestURL = "https://api.manager.wowma.jp/wmshopapi/searchTradeInfoListProc?shopId=64197720&totalCount=100&startDate={startDate}&endDate={endDate}"\
	.format(startDate=startDate,endDate=endDate)
	
	headers = {"Content-type": "application/x-www-form-urlencoded","Authorization": "Bearer " + au_api_key}
	req = requests.get(url=requestURL,headers=headers)
	soup = BeautifulSoup(req.text,"xml")
	orders = soup.find_all("orderInfo")
	
	#ステータスがキャンセルのものはあろうがなかろうが更新
	#ステータスが完了で既にあるデータは転記しない
	#ステータスが完了で未転記のデータは最下行に転記実行
	for i, order in enumerate(orders):
		
		OrderId = order.find("orderId").text
		orderDate = datetime.datetime.strptime(order.find("orderDate").text,"%Y/%m/%d %H:%M")
		OrderTime = orderDate.strftime("%Y/%m/%d")
			
		if not order.find("deliveryName").text == #処理対象の配送方法名を記入:
			continue
		
		if order.find("orderStatus").text == "完了":
			continue
		
		if order.find("orderStatus").text == "キャンセル":
			continue
		
		if OrderId in Order_ID_list:
			continue
		
		if OrderId in Memory_Order_list:
			continue
				
		
		details = order.find_all("detail")
		for z,detail in enumerate(details):

			product_code = detail.find("itemCode").text		
			Quantity = detail.find("unit").text
			price = detail.find("totalItemPrice").text
					
			if product_code[:3] == "set" or product_code[:3] == "SET":
				for set in SET_dic[product_code]:
					New_Order_list.append([set[0],set[1]*Quantity,OrderId])
			else:
				New_Order_list.append([product_code,Quantity,OrderId])
	
		#Qoo10
	T_now = datetime.datetime.now()
	T_before = T_now + datetime.timedelta(days=-30)
	
	endDate = T_now.strftime('%Y%m%d')
	startDate = T_before.strftime('%Y%m%d')
	requestURL = "https://api.qoo10.jp/GMKT.INC.Front.QAPIService/ebayjapan.qapi??v=1.0&returnType=xml&ShippingStat=2&method=ShippingBasic.GetShippingInfo_v2&key={Qoo10_SAK}&search_Sdate={startDate}&search_Edate={endDate}"\
	.format(Qoo10_SAK=Qoo10_SAK,startDate=startDate,endDate=endDate)
	req = requests.get(url=requestURL)
	soup = BeautifulSoup(req.text,"xml")
	orders = soup.find_all("ShippingInfo")
	
	for i, order in enumerate(orders):
		target = 7
		
		OrderId = order.find("orderNo").text
		orderDate = datetime.datetime.strptime(order.find("orderDate").text,"%Y-%m-%d %H:%M:%S")
		OrderTime = orderDate.strftime("%Y/%m/%d")
		
		if not order.find("DeliveryCompany").text == #処理対象の配送方法名を記入:
			continue
		
			
		product_code = order.find("sellerItemCode").text		
		Quantity = order.find("orderQty").text
		price = order.find("orderPrice").text.replace(".0000","")
			
		if product_code[:3] == "set" or product_code[:3] == "SET":
			for set in SET_dic[product_code]:
				New_Order_list.append([ProductSheetNo_To_ProductCode[set[0]],set[1]*Quantity,OrderId])
		else:
			New_Order_list.append([product_code,Quantity,OrderId])
	
		#在庫から注文を引く
	for cell in Memory_Order_lists:
		try:
			Inventory_dic[cell[0]] += int(cell[1]) * -1
		except:
			print("Inventory_dic[{}]：エラー".format(cell[0]))
	
		#在庫から注文を引く
	for cell in New_Order_list:
		try:
			Inventory_dic[cell[0]] += int(cell[1]) * -1
			if Inventory_dic[cell[0]] == 0:
				if not cell[0] in Do_ZERO_list:
					DO_ZERO_NEWS.append([cell[0]])
					ItemID = cell[0]
					quantity = 0
					au_stock_update(au_api_key,ItemID,quantity)
					yahoo_stock_update(access_token,ItemID,quantity)
					Qoo10UpdateStock(Qoo10_SAK,ItemID,quantity)
					print(cell[0],":在庫0にしました")
		except:
			print("Inventory_dic[{}]：エラー".format(cell[0]))
	write_csv(Memory_Order_path,New_Order_list)
	write_csv(DO_ZERO_Path,DO_ZERO_NEWS)
	print("Moniter_Finish")

au_api_key = os.getenv("au_api_key")
refresh_token_base = os.getenv("refresh_token_base")
Qoo10_SAK = os.getenv("Qoo10_SAK")
book_key = os.getenv("CosmeProduct")
sh_list = ["仕入れ商品リスト","出品商品リスト","仕入","引当","在庫","受注","新規注文","セット商品リスト"]

Inventory_Monitoring(book_key,sh_list,refresh_token_base,au_api_key)