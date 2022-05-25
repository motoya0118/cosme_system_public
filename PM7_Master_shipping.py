from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.select import Select
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

	credentials = ServiceAccountCredentials.from_json_keyfile_name('秘密鍵のJSON', scope)
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
	
def Yamato_shippinginfo_download():
	
	Today = datetime.datetime.now()
	InputDate = Today.strftime("%Y/%m/%d")
	iDir = os.getcwd() + "/入力/YamatoFF_追跡番号"
	options = Options()
	options = webdriver.ChromeOptions()
	options.add_experimental_option("prefs", {"download.default_directory": iDir })
	options.add_argument('--headless')
	options.add_argument("--no-sandbox")
	options.add_argument("--remote-debugging-port=9222")
	driver = webdriver.Chrome(options=options)
	driver = Yamato_login(driver)
	time.sleep(1)
	
	url = "https://kuroneko-ylc.com/ffportal/Ebina/shippings/"
	driver.get(url)
	time.sleep(5)
	flag = True
	while flag:
		try:
			ship_state = driver.find_element_by_xpath("//*[@id='__layout']/div/section/div/section/div[2]/main/section/div/div[2]/div[2]/table/tbody/tr[1]/td[3]").text
			print(ship_state)
			if not ship_state == "確定済":
				if not ship_state == "発送":
					flag = False
					return True
			
			flag = False
		except:print("ステータス確認失敗")
			
	flag = True
	while flag:
		url = "https://kuroneko-ylc.com/ffportal/Ebina/shippings/download"
		driver.get(url)
		time.sleep(10)
		try:	
			driver.find_element_by_xpath('//*[@id="__layout"]/div/section/div/section/div[2]/main/section/div/section/div[2]/div/div/div[1]/div/div/div[1]/div/input').send_keys(InputDate)
			element = driver.find_element_by_xpath('//*[@id="__layout"]/div/section/div/section/div[2]/main/section/div/section/div[2]/div/div/div[2]/div/span/select')
			text = "全て"
			select = Select(element)
			select.select_by_visible_text(text)
			driver.find_element_by_xpath("//*[@id='__layout']/div/section/div/section/div[2]/main/section/div/section/div[3]/p[2]/button").click()
			time.sleep(10)
			input_csv_path = glob_csv_path("./入力/YamatoFF_追跡番号")
			flag = False
			return False
		except Exception as e:
			print ('=== エラー内容 ===')
			print ('type:' + str(type(e)))
			print ('args:' + str(e.args))
			print ('e自身:' + str(e))

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


def create_message(sender, to, subject, message_text,store_name):
    message = MIMEText(message_text)
    message['to'] = to
    message['from'] = email.utils.formataddr((store_name, sender))
    message['subject'] = subject
    encode_message = base64.urlsafe_b64encode(message.as_bytes())
    return {'raw': encode_message.decode()}
# 3. メール送信の実行
def send_message(service, user_id, message):
    try:
        message = (service.users().messages().send(userId=user_id, body=message)
                   .execute())
        print('Message Id: %s' % message['id'])
        return message
    except errors.HttpError as error:
        print('An error occurred: %s' % error)
# 4. メインとなる処理

def send_gmail_Ship_Notification(BillFirstName,BillLastName,OrderId,OrderTime,PayMethodName,ShipMethodName,ShipRequestDate,ShipRequestTime,ShipLastName,ShipFirstName,ShipZipCode,ShipPrefecture,ShipCity,ShipAddress1,ShipAddress2,ShipPhoneNumber,Title,TotalPrice,UsePoint,ShipCharge,TotalCouponDiscount,SettleAmount,Total_cal,BillMailAddress,TotalMallCouponDiscount,store_name,store_adress,store_zip,store_url,store_mail,ShipInvoiceNumber):
	# 5. アクセストークンの取得
	SCOPES = ['https://www.googleapis.com/auth/gmail.send']
	creds = None
	if os.path.exists('token.pickle'):
		with open('token.pickle', 'rb') as token:
			creds = pickle.load(token)
	if not creds or not creds.valid:
		if creds and creds.expired and creds.refresh_token:
			creds.refresh(Request())
		else:
			flow = InstalledAppFlow.from_client_secrets_file(
			'credentials.json', SCOPES)
			creds = flow.run_local_server()
			with open('token.pickle', 'wb') as token:
				pickle.dump(creds, token)
	service = build('gmail', 'v1', credentials=creds)
    # 6. メール本文の作成
	sender = store_mail
	to = BillMailAddress
	subject = '【{store_name}】商品発送の御連絡'.format(store_name=store_name)
	message_text = """
----------------------------------------------------------------------
このメールはお客様の注文に関する大切なメールです。
お取引が完了するまで保存してください。
----------------------------------------------------------------------
{BillLastName} {BillFirstName} 様

 {store_name}：担当の村上と申します。

この度はご注文頂きまして誠にありがとうございます。

御注文の商品を本日発送致しましたので下記の通り御連絡致します。

配送方法：ヤマト配送
配送伝票番号：{ShipInvoiceNumber}
お荷物検索URL：https://toi.kuronekoyamato.co.jp/cgi-bin/tneko


下記の内容でご注文を承っております。

------------------------
------------------------
[注文ID]{OrderId}
[日時]{OrderTime}
[注文者] {BillLastName} {BillFirstName} 様
[支払方法] {PayMethodName}
[配送方法] {ShipMethodName}
[配送日時指定:]{ShipRequestDate} {ShipRequestTime}

--------------------------------
[送付先] {ShipLastName} {ShipFirstName} 様
      〒{ShipZipCode} {ShipPrefecture}{ShipCity}{ShipAddress1}{ShipAddress2}
      (TEL) {ShipPhoneNumber}
[商品]
      {Title}
****************************************************************
商品価格計   {Total_cal}円
※商品価格はストアクーポン値引き後の金額です
--------------------------------
小計        {Total_cal}円
ポイント利用 -{UsePoint}円
モールクーポン値引額 -{TotalMallCouponDiscount}円
送料       {ShipCharge}円
----------------------------------------------------------------
合計金額     {SettleAmount}円
----------------------------------------------------------------

このたびはご注文誠にありがとうございました。
今後とも『{store_name}』をどうぞよろしくお願い致します。

■□□□□□□□□□□□□□□□□□□□□■

【{store_name}】

【所在地】〒{store_zip}　
　　　　　{store_adress}
【ＵＲＬ】{store_url}
【メール】{store_mail}

■□□□□□□□□□□□□□□□□□□□□■
"""  .format(BillFirstName=BillFirstName,
			BillLastName=BillLastName,
			OrderId=OrderId,
			OrderTime=OrderTime,
			PayMethodName=PayMethodName,
			ShipMethodName=ShipMethodName,
			ShipRequestDate=ShipRequestDate,
			ShipRequestTime=ShipRequestTime,
			ShipLastName=ShipLastName,
			ShipFirstName=ShipFirstName,
			ShipZipCode=ShipZipCode,
			ShipPrefecture=ShipPrefecture,
			ShipCity=ShipCity,
			ShipAddress1=ShipAddress1,
			ShipAddress2=ShipAddress2,
			ShipPhoneNumber=ShipPhoneNumber,
			Title=Title,
			TotalPrice=TotalPrice,
			UsePoint=UsePoint,
			ShipCharge=ShipCharge,
			SettleAmount=SettleAmount,
			Total_cal=Total_cal,
			store_name=store_name,
			store_adress=store_adress,
			store_zip=store_zip,
			store_url=store_url,
			store_mail=store_mail,
			TotalMallCouponDiscount=TotalMallCouponDiscount,
			ShipInvoiceNumber=ShipInvoiceNumber) 
	message = create_message(sender, to, subject, message_text,store_name)
	# 7. Gmail APIを呼び出してメール送信
	send_message(service, 'me', message)

def send_gmail_after_Notification(BillFirstName,BillLastName,OrderId,OrderTime,PayMethodName,ShipMethodName,ShipRequestDate,ShipRequestTime,ShipLastName,ShipFirstName,ShipZipCode,ShipPrefecture,ShipCity,ShipAddress1,ShipAddress2,ShipPhoneNumber,Title,TotalPrice,UsePoint,ShipCharge,TotalCouponDiscount,SettleAmount,Total_cal,BillMailAddress,TotalMallCouponDiscount,store_name,store_adress,store_zip,store_url,store_mail,ShipInvoiceNumber):
	# 5. アクセストークンの取得
	SCOPES = ['https://www.googleapis.com/auth/gmail.send']
	creds = None
	if os.path.exists('token.pickle'):
		with open('token.pickle', 'rb') as token:
			creds = pickle.load(token)
	if not creds or not creds.valid:
		if creds and creds.expired and creds.refresh_token:
			creds.refresh(Request())
		else:
			flow = InstalledAppFlow.from_client_secrets_file(
			'credentials.json', SCOPES)
			creds = flow.run_local_server()
			with open('token.pickle', 'wb') as token:
				pickle.dump(creds, token)
	service = build('gmail', 'v1', credentials=creds)
    # 6. メール本文の作成
	sender = store_mail
	to = BillMailAddress
	subject = '【{store_name}】商品は無事お手元に届きましたでしょうか'.format(store_name=store_name)
	
	try:
		OrderTime = datetime.datetime.strptime(OrderTime,"%Y-%m-%d %H:%M")
	except:
		OrderTime = datetime.datetime.strptime(OrderTime,"%Y%m%d")
	
	OrderTime = OrderTime.strftime("%m{}%d{}").format(*"月日")
	
	message_text = """
<div style="text-align: center">
<img src="https://res.cloudinary.com/motoya/image/upload/v1644242457/%E8%A9%95%E4%BE%A1%E4%BE%9D%E9%A0%BC%E3%83%A1%E3%83%BC%E3%83%AB_mmk9hx.jpg"></img>
</div>
<br>
{BillLastName} {BillFirstName} 様<br>
<br>
 {store_name}：担当の村上と申します。<br>
<br>
先日は数あるショップの中から当店をご利用いただき<br>
誠にありがとうございました。<br>
<br>
スタッフ一同、心より感謝しております。<br>
<br>
{OrderTime}に御注文の商品は、無事お手元にお届けできましたでしょうか。<br>
<br>
私どもの商品が、少しでもお客様のお役に立てましたら大変嬉しく存じます。<br>
末永くお役立ていただけると幸いです。<br>
<br>
この度は御縁があり御注文頂きましたので<br>
【次回御注文3%off】クーポンを進呈させて頂きました。<br>
<br>
次回ご利用の際に当店の商品ページにクーポンが表示されますので<br>
是非次回御利用頂けますと幸いです。<br>
<br>
※クーポンの確認をする際に御注文履歴より当店の評価を頂けますと幸甚の至りでございます。<br>
<br>
------------------------<br>
御注文内容<br>
------------------------<br>
[注文ID]{OrderId}<br>
[日時]{OrderTime}<br>
[注文者] {BillLastName} {BillFirstName} 様<br>
[支払方法] {PayMethodName}<br>
[配送方法] {ShipMethodName}<br>
[配送日時指定:]{ShipRequestDate} {ShipRequestTime}<br>
<br>
--------------------------------<br>
[送付先] {ShipLastName} {ShipFirstName} 様<br>
      〒{ShipZipCode} {ShipPrefecture}{ShipCity}{ShipAddress1}{ShipAddress2}<br>
      (TEL) {ShipPhoneNumber}<br>
[商品]<br>
      {Title}<br>
****************************************************************<br>
商品価格計   {Total_cal}円<br>
※商品価格はストアクーポン値引き後の金額です<br>
--------------------------------<br>
小計        {Total_cal}円<br>
ポイント利用 -{UsePoint}円<br>
モールクーポン値引額 -{TotalMallCouponDiscount}円<br>
送料       {ShipCharge}円<br>
----------------------------------------------------------------<br>
合計金額     {SettleAmount}円<br>
----------------------------------------------------------------<br>
<br>
このたびはご注文誠にありがとうございました。<br>
今後とも『{store_name}』をどうぞよろしくお願い致します。<br>
<br>
■□□□□□□□□□□□□□□□□□□□□■<br>
<br>
【{store_name}】<br>
<br>
【所在地】〒{store_zip}　<br>
　　　　　{store_adress}<br>
【ＵＲＬ】{store_url}<br>
【メール】{store_mail}<br>
<br>
■□□□□□□□□□□□□□□□□□□□□■<br>
"""  .format(BillFirstName=BillFirstName,
			BillLastName=BillLastName,
			OrderId=OrderId,
			OrderTime=OrderTime,
			PayMethodName=PayMethodName,
			ShipMethodName=ShipMethodName,
			ShipRequestDate=ShipRequestDate,
			ShipRequestTime=ShipRequestTime,
			ShipLastName=ShipLastName,
			ShipFirstName=ShipFirstName,
			ShipZipCode=ShipZipCode,
			ShipPrefecture=ShipPrefecture,
			ShipCity=ShipCity,
			ShipAddress1=ShipAddress1,
			ShipAddress2=ShipAddress2,
			ShipPhoneNumber=ShipPhoneNumber,
			Title=Title,
			TotalPrice=TotalPrice,
			UsePoint=UsePoint,
			ShipCharge=ShipCharge,
			SettleAmount=SettleAmount,
			Total_cal=Total_cal,
			store_name=store_name,
			store_adress=store_adress,
			store_zip=store_zip,
			store_url=store_url,
			store_mail=store_mail,
			TotalMallCouponDiscount=TotalMallCouponDiscount,
			ShipInvoiceNumber=ShipInvoiceNumber) 
	message = create_message_html(sender, to, subject, message_text,store_name)
	# 7. Gmail APIを呼び出してメール送信
	send_message(service, 'me', message)

def create_message_html(sender, to, subject, message_text,store_name):
	message = MIMEText(message_text,"html")
	message['to'] = to
	message['from'] = email.utils.formataddr((store_name, sender))
	message['subject'] = subject
	message.add_header('Content-Type','text/html')
	encode_message = base64.urlsafe_b64encode(message.as_bytes())
	return {'raw': encode_message.decode()}


def au_Shipping_Notification(au_api_key,OrderId,ShipInvoiceNumber):
	
	#出荷ステータスの更新
	
	shippingDate = datetime.datetime.now()
	shippingDate = shippingDate.strftime("%Y/%m/%d")
	requestURL = "https://api.manager.wowma.jp/wmshopapi/updateTradeInfoProc?shopId=#ストアID(数字)"
	xmlPostBody = """<?xml version="1.0" encoding="UTF-8"?>
	<request>
		<shopId>#ストアID(数字)を記入</shopId>
		<orderId>{OrderId}</orderId>
		<shippingDate>{shippingDate}</shippingDate>
		<shippingCarrier>1</shippingCarrier>
		<shippingNumber>{ShipInvoiceNumber}</shippingNumber>
	</request>""".format(OrderId=OrderId,shippingDate=shippingDate,ShipInvoiceNumber=ShipInvoiceNumber)
	#ここにXML形式でデータを記入
	
	headers = {"Content-type": "application/xml;charset=UTF-8","Authorization": "Bearer " + au_api_key}
	
	# POSTリクエスト送信
	bytesXMLPostBody = xmlPostBody.encode("UTF-8")
	print(bytesXMLPostBody)
	req = requests.post(url=requestURL,data=bytesXMLPostBody,headers=headers)
	soup = BeautifulSoup(req.text, "xml")
	
	if not req.status_code == 200:
		print(soup)
		sys.exit("リフレッシュトークンを更新して下さい")
	
	#注文ステータスの更新
	requestURL = "https://api.manager.wowma.jp/wmshopapi/updateTradeStsProc"
	xmlPostBody = """<?xml version="1.0" encoding="UTF-8"?>
<request>
	<shopId>#ストアID(数字)を記入</shopId>
	<orderId>{OrderId}</orderId>
	<orderStatus>完了</orderStatus>
</request>""".format(OrderId=OrderId)
	#ここにXML形式でデータを記入
	
	headers = {"Content-type": "application/xml;charset=UTF-8","Authorization": "Bearer " + au_api_key}
	
	# POSTリクエスト送信
	bytesXMLPostBody = xmlPostBody.encode("utf-8")
	print(bytesXMLPostBody)
	req = requests.post(url=requestURL,data=bytesXMLPostBody,headers=headers)
	soup = BeautifulSoup(req.text, "xml")
	
	if not req.status_code == 200:
		print(soup)
		sys.exit("リフレッシュトークンを更新して下さい")

	#メールアドレスの取得
	requestURL = "https://api.manager.wowma.jp/wmshopapi/searchTradeInfoProc?shopId=#ストアID(数字)を記入&orderId={OrderId}"\
	.format(OrderId=OrderId)
	
	headers = {"Content-type": "application/x-www-form-urlencoded","Authorization": "Bearer " + au_api_key}
	req = requests.get(url=requestURL,headers=headers)
	soup = BeautifulSoup(req.text,"xml")
	
	order = soup
	
	details = order.find_all("detail")
	OrderId = order.find("orderId").text
	orderDate = datetime.datetime.strptime(order.find("orderDate").text,"%Y/%m/%d %H:%M")
	OrderTime = orderDate.strftime("%Y%m%d")
	TotalPrice = order.find("totalSalePrice").text
	ShipName = order.find("senderName").text
	ShipZipCode = order.find("senderZipCode").text
	shipAddresses = order.find("senderAddress").text.split(" ")
		
	shipAddress1 = ""
	shipAddress2 = ""
	for i,shipAddress in enumerate(shipAddresses):
		if i == 0:
			shipAddress1 = shipAddress
		elif i == 1:	
			shipAddress1 += shipAddress
		elif i == 2:
			shipAddress2 = shipAddress
		else:
			shipAddress2 += shipAddress
		
	ShipPhoneNumber = order.find("senderPhoneNumber1").text
	try:
		ShipRequestDate = order.find("deliveryRequestDay").text
	except:
		ShipRequestDate = None
		
	try:
		ShipRequestTime = order.find("deliveryRequestTime").text
	except:
		ShipRequestTime = None
	BillFirstName = ""
	BillLastName = order.find("ordererName").text
	PayMethodName = order.find("settlementName").text
	ShipMethodName = order.find("deliveryName").text
	ShipLastName = ShipName
	ShipFirstName = ""
	ShipPrefecture = ""
	ShipCity = ""
	ItemId = []
		
	for i,detail in enumerate(details):
		if i == 0:
			Title = detail.find("itemName").text + "　:　" + detail.find("itemPrice").text + "円　：　" + detail.find("unit").text + "個"
			Total_cal = detail.find("totalItemPrice").text
			ItemId.append(detail.find("itemCode").text)
		else:
			Title += "\n" + detail.find("itemName").text  + "　:　" + detail.find("itemPrice").text + "円　：　" + detail.find("unit").text + "個"
			Total_cal += detail.find("totalItemPrice").text
			ItemId.append(detail.find("itemCode").text) 
		
	ShipCharge = order.find("postagePrice").text
	TotalCouponDiscount = order.find("couponTotalPrice").text
	SettleAmount = order.find("requestPrice").text
	Total_cal = order.find("totalPrice").text
	BillMailAddress = order.find("mailAddress").text
	TotalMallCouponDiscount = TotalCouponDiscount

	Discount = order.find("discount").text
	UsePoint = order.find("useAuPoint").text
	Total = order.find("requestPrice").text
	Total_fee = round(int(Total)*0.1)
	ShipAddress1 = shipAddress1
	ShipAddress2 = shipAddress2
		
	#ストア情報
	store_name = #ストア名を記入
	store_adress = #ストア住所を記入
	store_zip = #ストア郵便番号を記入
	store_url = #ストアURLを記入
	store_mail = #メールアドレスを記入
	
	if not req.status_code == 200:
		sys.exit("リフレッシュトークンを更新して下さい")
	
	send_gmail_Ship_Notification(BillFirstName,BillLastName,OrderId,OrderTime,PayMethodName,ShipMethodName,ShipRequestDate,ShipRequestTime,ShipLastName,ShipFirstName,ShipZipCode,ShipPrefecture,ShipCity,ShipAddress1,ShipAddress2,ShipPhoneNumber,Title,TotalPrice,UsePoint,ShipCharge,TotalCouponDiscount,SettleAmount,Total_cal,BillMailAddress,TotalMallCouponDiscount,store_name,store_adress,store_zip,store_url,store_mail,ShipInvoiceNumber)


def yahoo_Shipping_Notification(refresh_token_base,OrderId,ShipInvoiceNumber):
	
	access_token = refresh_token(refresh_token_base)
	
	#出荷ステータスの更新
	url = "https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/orderShipStatusChange"
	headers = {"Authorization":"Bearer " + access_token}	
	ShipDate = datetime.datetime.now()
	ShipDate = ShipDate.strftime("%Y%m%d")
	
	data = """
<Req>
	<Target>
		<OrderId>{OrderId}</OrderId>
		<IsPointFix>true</IsPointFix>
	</Target>
	<Order>
		<Ship>
			<ShipStatus>3</ShipStatus>
			<ShipMethod>postage3</ShipMethod>
			<ShipCompanyCode>1001</ShipCompanyCode>
			<ShipInvoiceNumber1>{ShipInvoiceNumber}</ShipInvoiceNumber1>
			<ShipDate>{ShipDate}</ShipDate>
			<IsEazy>false</IsEazy>
		</Ship>
	</Order>
	<SellerId>chan-mu</SellerId>
</Req>""".format(OrderId=OrderId,ShipInvoiceNumber=ShipInvoiceNumber,ShipDate=ShipDate)
		
	req = requests.post(url=url,headers=headers,data=data,cert=(#crtファイルを指定,#keyファイルを指定))
	soup = BeautifulSoup(req.text, "xml")
	
	if not req.status_code == 200:
		print(soup)
		sys.exit("リフレッシュトークンを更新して下さい")
	
	#注文ステータスの更新
	url = "https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/orderStatusChange"
	headers = {"Authorization":"Bearer " + access_token}	
	
	data = """
<Req>
	<Target>
		<OrderId>{OrderId}</OrderId>
		<IsPointFix>true</IsPointFix>
	</Target>
	<Order>
		<OrderStatus>5</OrderStatus>
	</Order>
	<SellerId>chan-mu</SellerId>
</Req>""".format(OrderId=OrderId)
		
	req = requests.post(url=url,headers=headers,data=data,cert=(#crtファイルを指定,#keyファイルを指定))
	soup = BeautifulSoup(req.text, "xml")
	
	if not req.status_code == 200:
		print(soup)
		sys.exit("リフレッシュトークンを更新して下さい")

	#メールアドレスの取得
	url = "https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/orderInfo"
	headers = {"Authorization":"Bearer " + access_token}	
	
	data = """
		<Req>
			<Target>
			<OrderId>{OrderId}</OrderId>
			<Field>OrderId,ShipMethodName,OrderStatus,PayMethodName,PayStatus,Discount,BillMailAddress,BillFirstName,BillLastName,ItemId,UnitPrice,SubCode,OrderTime,Quantity,UnitPrice,TotalPrice,ShipFirstName,ShipLastName,ShipMethod,ShipZipCode,ShipPrefecture,ShipCity,ShipAddress1,ShipAddress2,ShipPhoneNumber,ShipRequestDate,ShipRequestTime,Title,Discount,UsePoint,TotalPrice,TotalCouponDiscount,SettleAmount,ShipCharge,TotalMallCouponDiscount</Field>
			</Target>
			<SellerId>chan-mu</SellerId>
		</Req>""".format(OrderId = OrderId)
		
	req = requests.post(url=url,headers=headers,data=data,cert=('./SHP-chan-mu/SHP-chan-mu.crt','./SHP-chan-mu/chan-mu.key'))
	soup = BeautifulSoup(req.text, "xml")
		
	BillFirstName = soup.find("BillFirstName").text
	BillLastName = soup.find("BillLastName").text
	OrderId = soup.find("OrderId").text
	OrderTime = soup.find("OrderTime").text
	OrderTime = datetime.datetime.fromisoformat(OrderTime)
	OrderTime = datetime.datetime.strftime(OrderTime,'%Y-%m-%d %H:%M')
	BillMailAddress = soup.find("BillMailAddress").text
	PayMethodName = soup.find("PayMethodName").text
	ShipMethodName = soup.find("ShipMethodName").text
	ShipRequestDate = soup.find("ShipRequestDate").text
	ShipRequestTime = soup.find("ShipRequestTime").text
	ShipLastName = soup.find("ShipLastName").text
	ShipFirstName = soup.find("ShipFirstName").text
	ShipZipCode = soup.find("ShipZipCode").text
	ShipPrefecture = soup.find("ShipPrefecture").text
	ShipCity = soup.find("ShipCity").text
	ShipAddress1 = soup.find("ShipAddress1").text
	ShipAddress2 = soup.find("ShipAddress2").text
	ShipPhoneNumber = soup.find("ShipPhoneNumber").text
	Discount = soup.find("Discount").text
	TotalMallCouponDiscount = soup.find("TotalMallCouponDiscount").text
	Items = soup.find_all("Item")
	Title = None
	ItemId = []
	for i,Item in enumerate(Items):
		if i == 0:
			Title = Item.find("Title").text + "　:　" + Item.find("UnitPrice").text + "円　：　" + Item.find("Quantity").text + "個"
			Total_cal = int(Item.find("UnitPrice").text) * int(Item.find("Quantity").text)
		else:
			Title += "\n" + Item.find("Title").text  + "　:　" + Item.find("UnitPrice").text + "円　：　" + Item.find("Quantity").text + "個"
			Total_cal += int(Item.find("UnitPrice").text) * int(Item.find("Quantity").text)			
	
	TotalPrice = soup.find("TotalPrice").text
	UsePoint = soup.find("UsePoint").text
	ShipCharge = soup.find("ShipCharge").text
	TotalCouponDiscount = soup.find("TotalCouponDiscount").text
	SettleAmount = soup.find("SettleAmount").text
		
	#ストア情報
	store_name = #ストア名を記入
	store_adress = #住所を記入
	store_zip = #郵便番号を記入
	store_url = #ストアURLを記入
	store_mail = #メールアドレスを記入
	
	if not req.status_code == 200:
		sys.exit("リフレッシュトークンを更新して下さい")
	
	send_gmail_Ship_Notification(BillFirstName,BillLastName,OrderId,OrderTime,PayMethodName,ShipMethodName,ShipRequestDate,ShipRequestTime,ShipLastName,ShipFirstName,ShipZipCode,ShipPrefecture,ShipCity,ShipAddress1,ShipAddress2,ShipPhoneNumber,Title,TotalPrice,UsePoint,ShipCharge,TotalCouponDiscount,SettleAmount,Total_cal,BillMailAddress,TotalMallCouponDiscount,store_name,store_adress,store_zip,store_url,store_mail,ShipInvoiceNumber)

def yahoo_after_Notification(refresh_token_base,OrderId,ShipInvoiceNumber):
	
	access_token = refresh_token(refresh_token_base)
	
	#メールアドレスの取得
	url = "https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/orderInfo"
	headers = {"Authorization":"Bearer " + access_token}	
	
	data = """
		<Req>
			<Target>
			<OrderId>{OrderId}</OrderId>
			<Field>OrderId,ShipMethodName,OrderStatus,PayMethodName,PayStatus,Discount,BillMailAddress,BillFirstName,BillLastName,ItemId,UnitPrice,SubCode,OrderTime,Quantity,UnitPrice,TotalPrice,ShipFirstName,ShipLastName,ShipMethod,ShipZipCode,ShipPrefecture,ShipCity,ShipAddress1,ShipAddress2,ShipPhoneNumber,ShipRequestDate,ShipRequestTime,Title,Discount,UsePoint,TotalPrice,TotalCouponDiscount,SettleAmount,ShipCharge,TotalMallCouponDiscount</Field>
			</Target>
			<SellerId>#SellerIdを記入</SellerId>
		</Req>""".format(OrderId = OrderId)
		
	req = requests.post(url=url,headers=headers,data=data,cert=(#crtファイルを指定,#keyファイルを指定))
	soup = BeautifulSoup(req.text, "xml")
		
	BillFirstName = soup.find("BillFirstName").text
	BillLastName = soup.find("BillLastName").text
	OrderId = soup.find("OrderId").text
	OrderTime = soup.find("OrderTime").text
	OrderTime = datetime.datetime.fromisoformat(OrderTime)
	OrderTime = datetime.datetime.strftime(OrderTime,'%Y-%m-%d %H:%M')
	BillMailAddress = soup.find("BillMailAddress").text
	PayMethodName = soup.find("PayMethodName").text
	ShipMethodName = soup.find("ShipMethodName").text
	ShipRequestDate = soup.find("ShipRequestDate").text
	ShipRequestTime = soup.find("ShipRequestTime").text
	ShipLastName = soup.find("ShipLastName").text
	ShipFirstName = soup.find("ShipFirstName").text
	ShipZipCode = soup.find("ShipZipCode").text
	ShipPrefecture = soup.find("ShipPrefecture").text
	ShipCity = soup.find("ShipCity").text
	ShipAddress1 = soup.find("ShipAddress1").text
	ShipAddress2 = soup.find("ShipAddress2").text
	ShipPhoneNumber = soup.find("ShipPhoneNumber").text
	Discount = soup.find("Discount").text
	TotalMallCouponDiscount = soup.find("TotalMallCouponDiscount").text
	Items = soup.find_all("Item")
	Title = None
	ItemId = []
	for i,Item in enumerate(Items):
		if i == 0:
			Title = soup.find("Title").text + "　:　" + soup.find("UnitPrice").text + "円　：　" + soup.find("Quantity").text + "個"
			Total_cal = int(soup.find("UnitPrice").text) * int(soup.find("Quantity").text)
			ItemId.append(soup.find("ItemId").text)
		else:
			Title += "\n" + soup.find("Title").text  + "　:　" + soup.find("UnitPrice").text + "円　：　" + soup.find("Quantity").text + "個"
			Total_cal += int(soup.find("UnitPrice").text) * int(soup.find("Quantity").text)
			ItemId.append(soup.find("ItemId").text)
				
	TotalPrice = soup.find("TotalPrice").text
	UsePoint = soup.find("UsePoint").text
	ShipCharge = soup.find("ShipCharge").text
	TotalCouponDiscount = soup.find("TotalCouponDiscount").text
	SettleAmount = soup.find("SettleAmount").text
		
	#ストア情報
	store_name = #ストア名を記入
	store_adress = #住所を記入
	store_zip = #郵便番号を記入
	store_url = #ストアURLを記入
	store_mail = #メールアドレスを記入
	
	if not req.status_code == 200:
		sys.exit("リフレッシュトークンを更新して下さい")
	
	send_gmail_after_Notification(BillFirstName,BillLastName,OrderId,OrderTime,PayMethodName,ShipMethodName,ShipRequestDate,ShipRequestTime,ShipLastName,ShipFirstName,ShipZipCode,ShipPrefecture,ShipCity,ShipAddress1,ShipAddress2,ShipPhoneNumber,Title,TotalPrice,UsePoint,ShipCharge,TotalCouponDiscount,SettleAmount,Total_cal,BillMailAddress,TotalMallCouponDiscount,store_name,store_adress,store_zip,store_url,store_mail,ShipInvoiceNumber)

def au_after_Notification(au_api_key,OrderId,ShipInvoiceNumber):
	
	#メールアドレスの取得
	requestURL = "https://api.manager.wowma.jp/wmshopapi/searchTradeInfoProc?shopId=#ストアID(数字)を入力&orderId={OrderId}"\
	.format(OrderId=OrderId)
	
	headers = {"Content-type": "application/x-www-form-urlencoded","Authorization": "Bearer " + au_api_key}
	req = requests.get(url=requestURL,headers=headers)
	soup = BeautifulSoup(req.text,"xml")
	
	order = soup
	
	details = order.find_all("detail")
	OrderId = order.find("orderId").text
	orderDate = datetime.datetime.strptime(order.find("orderDate").text,"%Y/%m/%d %H:%M")
	OrderTime = orderDate.strftime("%Y%m%d")
	TotalPrice = order.find("totalSalePrice").text
	ShipName = order.find("senderName").text
	ShipZipCode = order.find("senderZipCode").text
	shipAddresses = order.find("senderAddress").text.split(" ")
		
	shipAddress1 = ""
	shipAddress2 = ""
	for i,shipAddress in enumerate(shipAddresses):
		if i == 0:
			shipAddress1 = shipAddress
		elif i == 1:	
			shipAddress1 += shipAddress
		elif i == 2:
			shipAddress2 = shipAddress
		else:
			shipAddress2 += shipAddress
		
	ShipPhoneNumber = order.find("senderPhoneNumber1").text
	try:
		ShipRequestDate = order.find("deliveryRequestDay").text
	except:
		ShipRequestDate = None
		
	try:
		ShipRequestTime = order.find("deliveryRequestTime").text
	except:
		ShipRequestTime = None
	BillFirstName = ""
	BillLastName = order.find("ordererName").text
	PayMethodName = order.find("settlementName").text
	ShipMethodName = order.find("deliveryName").text
	ShipLastName = ShipName
	ShipFirstName = ""
	ShipPrefecture = ""
	ShipCity = ""
	ItemId = []
		
	for i,detail in enumerate(details):
		if i == 0:
			Title = detail.find("itemName").text + "　:　" + detail.find("itemPrice").text + "円　：　" + detail.find("unit").text + "個"
			Total_cal = detail.find("totalItemPrice").text
			ItemId.append(detail.find("itemCode").text)
		else:
			Title += "\n" + detail.find("itemName").text  + "　:　" + detail.find("itemPrice").text + "円　：　" + detail.find("unit").text + "個"
			Total_cal += detail.find("totalItemPrice").text
			ItemId.append(detail.find("itemCode").text) 
		
	ShipCharge = order.find("postagePrice").text
	TotalCouponDiscount = order.find("couponTotalPrice").text
	SettleAmount = order.find("requestPrice").text
	Total_cal = order.find("totalPrice").text
	BillMailAddress = order.find("mailAddress").text
	TotalMallCouponDiscount = TotalCouponDiscount

	Discount = order.find("discount").text
	UsePoint = order.find("useAuPoint").text
	Total = order.find("requestPrice").text
	Total_fee = round(int(Total)*0.1)
	ShipAddress1 = shipAddress1
	ShipAddress2 = shipAddress2
		
	#ストア情報
	store_name = #ストア名を記入
	store_adress = #住所を記入
	store_zip = #郵便番号を記入
	store_url = #ストアURLを記入
	store_mail = #メールアドレスを記入
	
	if not req.status_code == 200:
		sys.exit("リフレッシュトークンを更新して下さい")
	
	send_gmail_after_Notification(BillFirstName,BillLastName,OrderId,OrderTime,PayMethodName,ShipMethodName,ShipRequestDate,ShipRequestTime,ShipLastName,ShipFirstName,ShipZipCode,ShipPrefecture,ShipCity,ShipAddress1,ShipAddress2,ShipPhoneNumber,Title,TotalPrice,UsePoint,ShipCharge,TotalCouponDiscount,SettleAmount,Total_cal,BillMailAddress,TotalMallCouponDiscount,store_name,store_adress,store_zip,store_url,store_mail,ShipInvoiceNumber)

def refresh_token(refresh_token):
	
	app_id = os.getenv("app_id")
	secret = os.getenv("secret")
	
	headers = {"Accept":"*/*","Content-Type":"application/x-www-form-urlencoded; charset=utf-8"}
	data = "grant_type=refresh_token&client_id={app_id}&client_secret={secret}&refresh_token={refresh_token}".format(app_id=app_id,secret=secret,refresh_token=refresh_token)
	
	url = "https://auth.login.yahoo.co.jp/yconnect/v2/token"

	req = requests.post(url=url,headers=headers,data=data)
	token_directry = json.loads(req.text)
	return token_directry["access_token"]

def puchage_allocation(book_key,sh_list):
	work_book = gspread_me(book_key)

	work_sheet = work_book.worksheet(sh_list[2]) #仕入
	get_all = work_sheet.get_all_values()
	count = len(get_all)
	flag_cols = "K1:K" + str(count)
	flag_col = work_sheet.range(flag_cols)
	Allocations = []

	for i,row in enumerate(get_all):
		if i == 0:
			continue
	
		if len(row[10]) == 0:
			if len(row[9]) > 0:
				if len(row[0]) > 0:
					ID = row[0]
					product_name = row[1]
					capacity = row[2]
					quantity = row[5]
					price = int(float(row[4].replace(",","").replace("'",""))) / int(quantity)
					seller = row[6]
					seller_ID = row[7]
					Period_of_use = row[8]
					in_stock_date = row[9]
				
					flag_col[i].value = "TRUE"
				
					for n in range(int(quantity)):
						Allocation_ID = str(seller_ID) + "-" + str(n)
					
						Allocation = [ID,
						product_name,
						capacity,
						price,
						seller,
						seller_ID,
						Period_of_use,
						in_stock_date,
						Allocation_ID]
					
						Allocations.append(Allocation)

	work_sheet.update_cells(flag_col)

	work_sheet = work_book.worksheet(sh_list[3]) #引当
	get_all = work_sheet.get_all_values()

	count_allocation_sheet = len(get_all)
	count_allocations = len(Allocations)

	start_point = "A" + str(count_allocation_sheet + 1)
	end_point = "I" + str(count_allocation_sheet + 1 + count_allocations)
	set_key = start_point + ":" + end_point
	allocation_range = work_sheet.range(set_key)

	for n, row_data in enumerate(Allocations):
		for z, cell_data in enumerate(row_data):
			allocation_range[z + 9 * n].value = cell_data

	work_sheet.update_cells(allocation_range,value_input_option="USER_ENTERED")
	print("puchage_allocation_Finish")	

def order_organize_yahoo(book_key,sh_list,refresh_token_base):
	
	#gspread側
	work_book = gspread_me(book_key)
	Listing_sheet = work_book.worksheet(sh_list[1])
	Listing_all = Listing_sheet.get_all_values()
	count = len(Listing_all)
	Listing_flag_cols = "T1:T" + str(count)
	Listing_flag_col = Listing_sheet.range(Listing_flag_cols)
	
	Allocation_sheet = work_book.worksheet(sh_list[3])
	Allocation_all = Allocation_sheet.get_all_values()
	count = len(Allocation_all)
	Allocation_flag_cols = "J1:L" + str(count)
	Allocation_flag_col = Allocation_sheet.range(Allocation_flag_cols)
	order_sheet = work_book.worksheet(sh_list[5])
	order_all = order_sheet.get_all_values()
	count = len(order_all) + 300
	key = "A1:F" + str(count)
	order_range = order_sheet.range(key)
	
	#ID変換の辞書作成
	Directory_Listing_Code = {}
	Directory_Listing_Price = {}
	
	for row in Listing_all:
		if row[18] == "TRUE":
			Directory_Listing_Code[row[5]] = row[0]
			Directory_Listing_Price[row[0]] = row[10]
			
	#api側	
	refresh_token_base = refresh_token_base
		
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
		<Field>OrderId,ShipMethodName,OrderStatus,PayStatus,ItemId,ShipMethodName,TotalPrice</Field>
	</Search>
	<SellerId>chan-mu</SellerId>
	</Req>""".format(OrderTimeFrom=T_from,OrderTimeTo=T_to)
		
	req = requests.post(url=url,headers=headers,data=data,cert=(#crtファイルを指定,#keyファイルを指定))
	soup = BeautifulSoup(req.text, "xml")
	trees = []
	trees = soup.find_all("OrderInfo")

	
	if not req.status_code == 200:
		sys.exit("リフレッシュトークンを更新して下さい")
	
	#ステータスがキャンセルのものはあろうがなかろうが更新
	#ステータスが完了で既にあるデータは転記しない
	#ステータスが完了で未転記のデータは最下行に転記実行
	signal = True
	for i, data in enumerate(trees):
		target = 7
		
		if i == 0:
			continue
		
		amount_price = int(data.find("TotalPrice").text)
		OrderId = data.find("OrderId").text
		
		if not data.find("ShipMethodName").text == #対象の指定配送方法名を記入:
			continue
		
		if not data.find("OrderStatus").text == "5":
			if not data.find("OrderStatus").text == "4":
				continue
		
		for n, cell in enumerate(order_range):
			
			if n == target:
				target += 6
				if cell.value == OrderId:
					if data.find("OrderStatus").text == "4":
						order_range[n + 1].value = 0
						order_range[n + 2].value = 0
						
						if not cell.value == order_range[n + 6].value:
							break
					
					else:break
				
			
			if n % 6 == 0 and len(cell.value) == 0:
				print(OrderId)
				
				data = """
				<Req>
					<Target>
						<OrderId>{OrderId}</OrderId>
						<Field>OrderId,ShipMethodName,OrderStatus,PayMethodName,PayStatus,Discount,BillMailAddress,BillFirstName,BillLastName,ItemId,UnitPrice,SubCode,OrderTime,Quantity,UnitPrice,TotalPrice,ShipFirstName,ShipLastName,ShipMethod,ShipZipCode,ShipPrefecture,ShipCity,ShipAddress1,ShipAddress2,ShipPhoneNumber,ShipRequestDate,ShipRequestTime,Title,Discount,UsePoint,TotalPrice,TotalCouponDiscount,SettleAmount,ShipCharge,TotalMallCouponDiscount</Field>
					</Target>
					<SellerId>chan-mu</SellerId>
				</Req>""".format(OrderId = OrderId)
				
				url = "https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/orderInfo"
				headers = {"Authorization":"Bearer " + access_token}
				req = requests.post(url=url,headers=headers,data=data,cert=(#crtファイルを指定,#keyファイルを指定))
				soup = BeautifulSoup(req.text, "xml")
				
				OrderId = soup.find("OrderId").text
				OrderTime = soup.find("OrderTime").text
				OrderTime = datetime.datetime.fromisoformat(OrderTime)
				Items = soup.find_all("Item")
				
				#何分割するか確認
				#出品IDを確認
				ct = 0
				for z,Item in enumerate(Items):					
					product_code = Item.find("ItemId").text
					quantity = Item.find("Quantity").text
					price = Item.find("UnitPrice").text
					
					sum_price = 0
					if product_code[:3] == "set" or product_code[:3] == "SET":
						Purchage_Item_List = Change_SetID_to_ListingID(work_book,sh_list,Directory_Listing_Code,Directory_Listing_Price,product_code,quantity)
						for row in Purchage_Item_List:
							sum_price += int(row[2])
					else:
						try:
							Purchage_Item_List = [[Directory_Listing_Code[product_code],quantity,price]]
							sum_price = int(price)
						except:
							continue
					
					for cell in Purchage_Item_List:
					#転記実行--受注シート
						order_range[n + 6*ct].value = cell[0]
						order_range[n+1 + 6*ct].value = OrderId
						order_range[n+2 + 6*ct].value = cell[1]
						order_range[n+3 + 6*ct].value = round(int(price) * int(quantity) * int(cell[2]) / sum_price)
						order_range[n+4 + 6*ct].value = OrderTime.strftime("%Y/%m/%d")
						ct += 1
				break
					
	order_sheet.update_cells(order_range,value_input_option="USER_ENTERED")
	print("order_organize_yahoo_Finish")


def order_organize_au(book_key,sh_list,au_api_key):
	
	#gspread側
	work_book = gspread_me(book_key)
	Listing_sheet = work_book.worksheet(sh_list[1])
	Listing_all = Listing_sheet.get_all_values()
	count = len(Listing_all)
	Listing_flag_cols = "T1:T" + str(count)
	Listing_flag_col = Listing_sheet.range(Listing_flag_cols)
	
	Allocation_sheet = work_book.worksheet(sh_list[3])
	Allocation_all = Allocation_sheet.get_all_values()
	count = len(Allocation_all)
	Allocation_flag_cols = "J1:L" + str(count)
	Allocation_flag_col = Allocation_sheet.range(Allocation_flag_cols)
		
	order_sheet = work_book.worksheet(sh_list[5])
	order_all = order_sheet.get_all_values()
	count = len(order_all) + 300
	key = "A1:F" + str(count)
	order_range = order_sheet.range(key)
	
	#ID変換の辞書作成
	Directory_Listing_Code = {}
	Directory_Listing_Price = {}
	
	for row in Listing_all:
		if row[18] == "TRUE":
			Directory_Listing_Code[row[5]] = row[0]
			Directory_Listing_Price[row[0]] = row[10]
	
	#api側
	au_url = "https://api.manager.wowma.jp/wmshopapi/"
	T_now = datetime.datetime.now()
	T_before = T_now + datetime.timedelta(days=-30)
	
	endDate = T_now.strftime('%Y%m%d')
	startDate = T_before.strftime('%Y%m%d')
	requestURL = "https://api.manager.wowma.jp/wmshopapi/searchTradeInfoListProc?shopId=#ストアID(数字)を入力&totalCount=100&startDate={startDate}&endDate={endDate}"\
	.format(startDate=startDate,endDate=endDate)
	
	headers = {"Content-type": "application/x-www-form-urlencoded","Authorization": "Bearer " + au_api_key}
	req = requests.get(url=requestURL,headers=headers)
	soup = BeautifulSoup(req.text,"xml")
	orders = soup.find_all("orderInfo")
	
	#ステータスがキャンセルのものはあろうがなかろうが更新
	#ステータスが完了で既にあるデータは転記しない
	#ステータスが完了で未転記のデータは最下行に転記実行
	for i, order in enumerate(orders):
		target = 7
		
		OrderId = order.find("orderId").text
		orderDate = datetime.datetime.strptime(order.find("orderDate").text,"%Y/%m/%d %H:%M")
		OrderTime = orderDate.strftime("%Y/%m/%d")
		
		
		if not order.find("deliveryName").text == #対象の指定配送名を記入:
			continue
		
		if not order.find("orderStatus").text == "完了":
			if not order.find("orderStatus").text == "キャンセル":
				continue
		
		for n, cell in enumerate(order_range):
			
			if n == target:
				target += 6
				if cell.value == OrderId:
					if order.find("orderStatus").text == "キャンセル":
						order_range[n + 1].value = 0
						order_range[n + 2].value = 0
						
						if not cell.value == order_range[n + 6].value:
							break
					
					elif len(order_range[n + 4].value) > 0:
						break
				
			
			if n % 6 == 0 and len(cell.value) == 0:
				
				ct = 0
				#出品IDを確認
				details = order.find_all("detail")
				for z,detail in enumerate(details):

					product_code = detail.find("itemCode").text		
					quantity = detail.find("unit").text
					price = detail.find("totalItemPrice").text
					
					sum_price = 0
					if product_code[:3] == "set" or product_code[:3] == "SET":
						Purchage_Item_List = Change_SetID_to_ListingID(work_book,sh_list,Directory_Listing_Code,Directory_Listing_Price,product_code,quantity)
						for row in Purchage_Item_List:
							sum_price += int(row[2])
					else:
						Purchage_Item_List = [[Directory_Listing_Code[product_code],quantity,price]]
						sum_price = int(price)
							
							
					#転記実行--受注シート
					for cell in Purchage_Item_List:
					#転記実行--受注シート
						order_range[n + 6*ct].value = cell[0]
						order_range[n+1 + 6*ct].value = OrderId
						order_range[n+2 + 6*ct].value = cell[1]
						order_range[n+3 + 6*ct].value = round(int(price) * int(quantity) * int(cell[2]) / sum_price)
						order_range[n+4 + 6*ct].value = OrderTime
						ct += 1
					
					#転記実行--引当シート
							
				break
					
	order_sheet.update_cells(order_range,value_input_option="USER_ENTERED")
	print("order_organize_au_Finish")

def order_organize_Qoo10(book_key,sh_list,Qoo10_SAK):
	
	#gspread側
	work_book = gspread_me(book_key)
	Listing_sheet = work_book.worksheet(sh_list[1])
	Listing_all = Listing_sheet.get_all_values()
	count = len(Listing_all)
	Listing_flag_cols = "T1:T" + str(count)
	Listing_flag_col = Listing_sheet.range(Listing_flag_cols)
	
	Allocation_sheet = work_book.worksheet(sh_list[3])
	Allocation_all = Allocation_sheet.get_all_values()
	count = len(Allocation_all)
	Allocation_flag_cols = "J1:L" + str(count)
	Allocation_flag_col = Allocation_sheet.range(Allocation_flag_cols)
		
	order_sheet = work_book.worksheet(sh_list[5])
	order_all = order_sheet.get_all_values()
	count = len(order_all) + 300
	key = "A1:F" + str(count)
	order_range = order_sheet.range(key)
	
	#ID変換の辞書作成
	Directory_Listing_Code = {}
	Directory_Listing_Price = {}
	
	for row in Listing_all:
		if row[18] == "TRUE":
			Directory_Listing_Code[row[5]] = row[0]
			Directory_Listing_Price[row[0]] = row[10]
	
	#api側
	T_now = datetime.datetime.now()
	T_before = T_now + datetime.timedelta(days=-30)
	
	endDate = T_now.strftime('%Y%m%d')
	startDate = T_before.strftime('%Y%m%d')
	requestURL = "https://api.qoo10.jp/GMKT.INC.Front.QAPIService/ebayjapan.qapi??v=1.0&returnType=xml&ShippingStat=4&method=ShippingBasic.GetShippingInfo_v2&key={Qoo10_SAK}&search_Sdate={startDate}&search_Edate={endDate}"\
	.format(Qoo10_SAK=Qoo10_SAK,startDate=startDate,endDate=endDate)
	req = requests.get(url=requestURL)
	soup = BeautifulSoup(req.text,"xml")
	orders = soup.find_all("ShippingInfo")
	
	#ステータスがキャンセルのものはあろうがなかろうが更新
	#ステータスが完了で既にあるデータは転記しない
	#ステータスが完了で未転記のデータは最下行に転記実行
	for i, order in enumerate(orders):
		target = 7
		
		OrderId = order.find("orderNo").text
		orderDate = datetime.datetime.strptime(order.find("orderDate").text,"%Y-%m-%d %H:%M:%S")
		OrderTime = orderDate.strftime("%Y/%m/%d")
		
		
		if not order.find("DeliveryCompany").text == #対象の指定配送方法名を記入:
			continue
		
		for n, cell in enumerate(order_range):
			
			if n == target:
				target += 6
				
			if cell.value == OrderId:
				break
			
			if n % 6 == 0 and len(cell.value) == 0:
				
				ct = 0
				
				product_code = order.find("sellerItemCode").text		
				quantity = order.find("orderQty").text
				price = order.find("orderPrice").text.replace(".0000","")
					
				sum_price = 0
				if product_code[:3] == "set" or product_code[:3] == "SET":
					Purchage_Item_List = Change_SetID_to_ListingID(work_book,sh_list,Directory_Listing_Code,Directory_Listing_Price,product_code,quantity)
					for row in Purchage_Item_List:
						sum_price += int(row[2])
				else:
					Purchage_Item_List = [[Directory_Listing_Code[product_code],quantity,price]]
					sum_price = int(price)
							
							
				#転記実行--受注シート
				for cell in Purchage_Item_List:
				#転記実行--受注シート
					order_range[n + 6*ct].value = cell[0]
					order_range[n+1 + 6*ct].value = OrderId
					order_range[n+2 + 6*ct].value = cell[1]
					order_range[n+3 + 6*ct].value = round(int(price) * int(quantity) * int(cell[2]) / sum_price)
					order_range[n+4 + 6*ct].value = OrderTime
					ct += 1
					
				#転記実行--引当シート
							
				break
					
	order_sheet.update_cells(order_range,value_input_option="USER_ENTERED")
	print("order_organize_Qoo10_Finish")

def Order_Allocation(book_key,sh_list):

	work_book = gspread_me(book_key)
	
	Allocation_sheet = work_book.worksheet(sh_list[3])
	Allocation_all = Allocation_sheet.get_all_values()
	count = len(Allocation_all)
	Allocation_key = "J1:L" + str(count)
	Allocation_ranges = Allocation_sheet.range(Allocation_key)
	
	order_sheet = work_book.worksheet(sh_list[5])
	order_all = order_sheet.get_all_values()
	count = len(order_all)
	order_key = "F1:F" + str(count)
	order_cols = order_sheet.range(order_key)
	
	Listing_sheet = work_book.worksheet(sh_list[1])
	Listing_getall = Listing_sheet.get_all_values()

	directry = {}

	for L_row in Listing_getall:
		if len(L_row[3]) > 0:
			directry[L_row[0]] = [[L_row[1],L_row[2]],[L_row[3],L_row[4]]]
		else:
			directry[L_row[0]] = [[L_row[1],L_row[2]]]

	for i,Order_row in enumerate(order_all):
		if i == 0:
			continue
		
		if len(Order_row[5]) > 0:
			continue
		
		Listing_ID = Order_row[0]
		Order_ID = Order_row[1]
		Quantity = int(Order_row[2])
		if Quantity == 0:
			Order_Price = 0
		else:
			Order_Price = int(Order_row[3])/Quantity
		Order_Date = Order_row[4]
		flag = True
		ct = 0
		
		for Quantity_Repeat in range(Quantity):
			for Listing_ID_Directry in directry[Listing_ID]:
				Purchage_ID = Listing_ID_Directry[0]
				Purchage_Quantity = int(Listing_ID_Directry[1].replace("'",""))
				for Allocation_Repeat in range(Purchage_Quantity):
					for n,Allocation_row in enumerate(Allocation_all):
						ct += 1
						if Allocation_row[0] == Purchage_ID:
							if len(Allocation_ranges[n*3 + 2].value) == 0:
								Allocation_ranges[n*3].value = Order_ID
								
								if len(Order_ID) == 16:
									Allocation_ranges[n*3 + 1].value = Order_Price/Purchage_Quantity * 0.92
								elif len(Order_ID) == 9:
									Allocation_ranges[n*3 + 1].value = Order_Price/Purchage_Quantity * 0.80
					
								Allocation_ranges[n*3 + 2].value = Order_Date
								break
							
							if n+1 == len(Allocation_all):
								flag = False
		
		if flag == False:
			print("下記の商品番号で引当未転記が発生したため処理を中止します")
			print("Order_ID:" + Order_ID)
			return None
			
		order_cols[i].value = "TRUE"
		print("Writed_OrderID:>>>" + Order_ID)
	
	order_sheet.update_cells(order_cols,value_input_option="USER_ENTERED")
	Allocation_sheet.update_cells(Allocation_ranges,value_input_option="USER_ENTERED")
	print("Order_Allocation_Finish")
	
def stock_allocation(book_key,sh_list):
	work_book = gspread_me(book_key)
	purchage_sheet = work_book.worksheet(sh_list[0])
	purchage_all = purchage_sheet.get_all_values()
	
	Allocation_sheet = work_book.worksheet(sh_list[3])
	Allocation_all = Allocation_sheet.get_all_values()
	
	Stock_sheet = work_book.worksheet(sh_list[4])
	Stock_all = Stock_sheet.get_all_values()
	count = len(purchage_all)
	Stock_key = "A1:L" + str(count)
	Stock_range = Stock_sheet.range(Stock_key)
	purchage_ID_dic = {}
	Purchage_ID_list = []
	
	for i,cell in enumerate(Stock_range):
		if i >11:
			cell.value = ""
	
	for i,purchage_cell in enumerate(purchage_all):
		if i == 0:
			continue
		
		Purchage_ID_list.append(purchage_cell[0])
		
		purchage_ID_dic[purchage_cell[0]] = {
		"Name":purchage_cell[1],
		"Capacity":purchage_cell[2],
		"Sell_count":0,
		"Stock":0,
		"Sell_value":0,
		"Purchage_value":0,
		"Date_Purchage":0,
		"Date_order":0,
		"Date_different":0,
		"purchage_flag":purchage_cell[3]}
	
	for n,Allocation_cell in enumerate(Allocation_all):
		if n == 0:
			continue
		
		if len(Allocation_cell[11]) == 0:
			purchage_ID_dic[Allocation_cell[0]]["Stock"] += 1
		
		else:
			purchage_ID_dic[Allocation_cell[0]]["Sell_count"] += 1
			purchage_ID_dic[Allocation_cell[0]]["Sell_value"] += float(Allocation_cell[10].replace(",",""))
			purchage_ID_dic[Allocation_cell[0]]["Purchage_value"] += float(Allocation_cell[3].replace(",",""))
		
		Date_Purchage = Allocation_cell[7].replace("'","").split("/")
		print(n)
		Date_Purchage = datetime.date(int(Date_Purchage[0]),int(Date_Purchage[1]),int(Date_Purchage[2]))
		
		try:
			if purchage_ID_dic[Allocation_cell[0]]["Date_order"] > Date_Purchage:
				Date_Purchage = purchage_ID_dic[Allocation_cell[0]]["Date_order"]
		except: purchage_ID_dic[Allocation_cell[0]]["Date_Purchage"] = Date_Purchage
		
		
		if len(Allocation_cell[11]) > 0:
			Date_order = Allocation_cell[11].replace("'","").split("/")
			Date_order = datetime.date(int(Date_order[0]),int(Date_order[1]),int(Date_order[2]))
			purchage_ID_dic[Allocation_cell[0]]["Date_order"] = Date_order
			purchage_ID_dic[Allocation_cell[0]]["Date_Purchage"] = Date_Purchage
			Date_different = Date_order - Date_Purchage
			purchage_ID_dic[Allocation_cell[0]]["Date_different"] += Date_different.days
		
	for z,Purchage_ID in enumerate(Purchage_ID_list):
		
		product_code = Purchage_ID
		product_Name = purchage_ID_dic[product_code]["Name"]
		Capacity = purchage_ID_dic[product_code]["Capacity"]
		Sell_count = purchage_ID_dic[product_code]["Sell_count"]
		Stock = purchage_ID_dic[product_code]["Stock"]
		Parchage_Flag = purchage_ID_dic[product_code]["purchage_flag"]
		
		if Sell_count > 0:
			Sell_Probability = 30 / (purchage_ID_dic[product_code]["Date_different"] / Sell_count)
			Standard_Stock = Sell_Probability * (30/30)
			Average_Benefit = purchage_ID_dic[product_code]["Sell_value"] / purchage_ID_dic[product_code]["Purchage_value"]
			target_price = purchage_ID_dic[product_code]["Sell_value"] /  Sell_count * 0.88
		else:
			Sell_Probability = 0
			Standard_Stock = 0
			Average_Benefit = 0
			target_price = 0
		
		
		Latest_Purchage = purchage_ID_dic[product_code]["Date_Purchage"]
		Latest_Order = purchage_ID_dic[product_code]["Date_order"]
		
		Stock_range[(z+1)*12].value = product_code
		Stock_range[(z+1)*12+1].value = product_Name
		Stock_range[(z+1)*12+2].value = Capacity
		Stock_range[(z+1)*12+3].value = Sell_count
		Stock_range[(z+1)*12+4].value = Stock
		Stock_range[(z+1)*12+5].value = round(Standard_Stock,1)
		Stock_range[(z+1)*12+6].value = round(Sell_Probability,2)
		Stock_range[(z+1)*12+7].value = round(Average_Benefit,2)
		Stock_range[(z+1)*12+8].value = round(target_price,1)
		Stock_range[(z+1)*12+11].value = Parchage_Flag
		if not Latest_Purchage == 0:
			Stock_range[(z+1)*12+9].value = Latest_Purchage.strftime("%Y/%m/%d")
		if not Latest_Order == 0:
			Stock_range[(z+1)*12+10].value = Latest_Order.strftime("%Y/%m/%d")
	
	Stock_sheet.update_cells(Stock_range,value_input_option="USER_ENTERED")
	print("stock_allocation_Finish")

def Qoo10_Shipping_Notification(Qoo10_SAK,ShipInvoiceNumber,Qoo10TargetList):
	
	#出荷ステータスの更新
	for row in Qoo10TargetList:
		OrderId = row[0]
		requestURL = "https://api.qoo10.jp/GMKT.INC.Front.QAPIService/ebayjapan.qapi"
		params = {"v":"1.0",
		"returnType":"xml",
		"method":"ShippingBasic.SetSendingInfo",
		"key":Qoo10_SAK,
		"OrderNo":OrderId,
		"ShippingCorp":#指定配送方法名を記入,
		"TrackingNo":ShipInvoiceNumber}
	
		req = requests.get(url=requestURL,params=params)
		soup = BeautifulSoup(req.text, "xml")
		print(soup)
	
		if not req.status_code == 200:
			print(soup)
			sys.exit("リフレッシュトークンを更新して下さい")

def Master_Shipping_Notification(book_key,sh_list,au_api_key,refresh_token_base,Qoo10_SAK):
	
	flag = True
	while flag:
		flag = Yamato_shippinginfo_download()
		if flag:
			time.sleep(3600)
	
	work_book = gspread_me(book_key)
	input_csv_path = glob_csv_path("./入力/YamatoFF_追跡番号")
	input_csvs = []
	input_csvs = read_csv(input_csv_path)
	
	New_Order_sheet = work_book.worksheet(sh_list[6])
	New_Order_all = New_Order_sheet.get_all_values()
	Maxrow = len(New_Order_all)
	count = Maxrow
	key = "A1:E" + str(count)
	order_range = New_Order_sheet.range(key)
	time_now = datetime.datetime.now()
	time_now = time_now.strftime("%Y/%m/%d")
	
	#Qoo10用のリスト生成
	Qoo10Dic = {}
	for n,New_Order in enumerate(New_Order_all):
			if len(New_Order[2]) == 0:
				if New_Order[1] == "Qoo10":
					Qoo10List = New_Order[0].split(",")
					if Qoo10List[0] in Qoo10Dic:
						Qoo10Dic[Qoo10List[0]].append([Qoo10List[1],n])
					else:
						Qoo10Dic[Qoo10List[0]] = [[Qoo10List[1],n]]
	
	for i,csv in enumerate(input_csvs):
		if i == 0:
			continue
		
		OrderId = csv[0]
		ShipInvoiceNumber = csv[1]		
		target_row = 0
		
		if OrderId[0] == "Q":
			#Qoo10の分岐
			CartNo = OrderId.replace("Q","")
			if not CartNo in Qoo10Dic:
				print(OrderId,"新規注文シートにないため未処理です")
				continue
			
			Qoo10TargetList = Qoo10Dic[CartNo]
		
		else:
			#ヤフショ・au
			for n,New_Order in enumerate(New_Order_all):
				if New_Order[0] == OrderId:
					if len(New_Order[2]) == 0:
						target_row = n
						break
			
			if target_row == 0:
				print(OrderId,"新規注文シートにないため未処理です")
				continue
	
		if len(OrderId) == 16:
			yahoo_Shipping_Notification(refresh_token_base,OrderId,ShipInvoiceNumber)
			time.sleep(1)
		elif OrderId[0] == "Q":
			Qoo10_Shipping_Notification(Qoo10_SAK,ShipInvoiceNumber,Qoo10TargetList)
			time.sleep(1)
		elif len(OrderId) == 9:
			au_Shipping_Notification(au_api_key,OrderId,ShipInvoiceNumber)
			time.sleep(1)
		else:
			print(OrderId,"未登録市場の注文番号判定")
		
		if OrderId[0] == "Q":
			for row in Qoo10TargetList:
				target_row = row[1]
				order_range[target_row * 5 + 2].value = "済"
				order_range[target_row * 5 + 3].value = time_now
				order_range[target_row * 5 + 4].value = time_now
		else:
			order_range[target_row * 5 + 2].value = "済"
			order_range[target_row * 5 + 3].value = time_now
	
	shutil.move(input_csv_path,"./入力/YamatoFF_追跡番号/処理済")
	New_Order_sheet.update_cells(order_range)
	print("Master_Shipping_Notification_Finish")
	
	puchage_allocation(book_key,sh_list)
	time.sleep(3)
	order_organize_yahoo(book_key,sh_list,refresh_token_base)
	time.sleep(3)
	order_organize_au(book_key,sh_list,au_api_key)
	time.sleep(3)
	order_organize_Qoo10(book_key,sh_list,Qoo10_SAK)
	time.sleep(3)
	Order_Allocation(book_key,sh_list)
	time.sleep(3)
	stock_allocation(book_key,sh_list)
	time.sleep(3)
	Master_after_Notification(book_key,sh_list,au_api_key,refresh_token_base)


	
def Master_after_Notification(book_key,sh_list,au_api_key,refresh_token_base):
	
	work_book = gspread_me(book_key)
	New_Order_sheet = work_book.worksheet(sh_list[6])
	New_Order_all = New_Order_sheet.get_all_values()
	Maxrow = len(New_Order_all)
	count = Maxrow
	key = "A1:E" + str(count)
	order_range = New_Order_sheet.range(key)
	time_now = datetime.datetime.now()
	time_now_str = time_now.strftime("%Y/%m/%d")
	
	for i,row in enumerate(New_Order_all):
		if i == 0:
			continue
		
		OrderId = row[0]
		Place = row[1]
		ShipInvoiceNumber = "after"
		
		if len(row[4]) > 0:
			continue
		
		if len(row[3]) == 0:
			continue
		
		shipping_date = datetime.datetime.strptime(row[3],"%Y/%m/%d")
		time_was = time_now + datetime.timedelta(days=-3)
		
		if shipping_date > time_was:
			continue
		
		if Place == "yahoo":
			yahoo_after_Notification(refresh_token_base,OrderId,ShipInvoiceNumber)
		elif Place == "au":
			au_after_Notification(au_api_key,OrderId,ShipInvoiceNumber)
		
		order_range[i * 5 + 4].value = time_now_str
	
	New_Order_sheet.update_cells(order_range)
	print("Master_Shipping_Notification_Finish")

au_api_key = os.getenv("au_api_key")
refresh_token_base = os.getenv("refresh_token_base")
Qoo10_SAK = os.getenv("Qoo10_SAK")
book_key = os.getenv("CosmeProduct")
sh_list = ["仕入れ商品リスト","出品商品リスト","仕入","引当","在庫","受注","新規注文","セット商品リスト"]

Master_Shipping_Notification(book_key,sh_list,au_api_key,refresh_token_base,Qoo10_SAK)