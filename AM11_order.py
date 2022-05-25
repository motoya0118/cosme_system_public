from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
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

def Yamato_shippinginfo_upload(upload_file):
	options = Options()
	options.add_argument('--headless')
	options.add_argument("--no-sandbox")
	options.add_argument("--remote-debugging-port=9222")
	driver = webdriver.Chrome(ChromeDriverManager().install(),options=options)
	Yamato_login(driver)
	#商品情報アップロード
	url = "https://kuroneko-ylc.com/ffportal/Ebina/shippings/upload"
	driver.get(url)
	try:
		time.sleep(10)
		driver.find_element_by_xpath("//*[@id='__layout']/div/section/div/section/div[2]/main/section/div/div[1]/label/input").send_keys(upload_file)
		driver.find_element_by_xpath("//*[@id='__layout']/div/section/div/section/div[2]/main/section/div/div[3]/p/button").click()
		time.sleep(10)
		driver.find_element_by_xpath("/html/body/div[2]/div[2]/footer/button[2]").click()
		time.sleep(10)
		return False
	except:
		return True

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

def send_gmail_first(BillFirstName,BillLastName,OrderId,OrderTime,PayMethodName,ShipMethodName,ShipRequestDate,ShipRequestTime,ShipLastName,ShipFirstName,ShipZipCode,ShipPrefecture,ShipCity,ShipAddress1,ShipAddress2,ShipPhoneNumber,Title,TotalPrice,UsePoint,ShipCharge,TotalCouponDiscount,SettleAmount,Total_cal,BillMailAddress,TotalMallCouponDiscount,store_name,store_adress,store_zip,store_url,store_mail):
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
	subject = '【{store_name}】御注文ありがとうございます'.format(store_name=store_name)
	message_text = """
----------------------------------------------------------------------
このメールはお客様の注文に関する大切なメールです。
お取引が完了するまで保存してください。
----------------------------------------------------------------------
{BillLastName} {BillFirstName} 様

 {store_name}：担当の村上と申します。

この度はご注文頂きまして誠にありがとうございます。

下記の内容でご注文を承りましたので、ご確認をお願い致します。

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

◆◇◆　ポスト投函配送についてご確認ください　◆◇◆
ネコポス（ポスト投函）につきましては、配送方法をご選択いただきましても、同梱の組み合わせが対応サイズを超える場合は宅配便にてお届けとなります。
あらかじめご了承ください。
------------------------
------------------------

------------------------
  商品のお届けについて
------------------------
当日11時までの御注文の場合：当日発送
当日11時以降の御注文の場合：翌日発送
------------------------

------------------------

商品の発送が完了致しましたら、運送会社・伝票番号等をメールにてご連絡致します。
恐れ入りますが、次のご案内まで今しばらくお待ち下さい。

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
			TotalMallCouponDiscount=TotalMallCouponDiscount) 
	message = create_message(sender, to, subject, message_text,store_name)
	# 7. Gmail APIを呼び出してメール送信
	send_message(service, 'me', message)

def send_gmail_delay(BillFirstName,BillLastName,OrderId,OrderTime,PayMethodName,ShipMethodName,ShipRequestDate,ShipRequestTime,ShipLastName,ShipFirstName,ShipZipCode,ShipPrefecture,ShipCity,ShipAddress1,ShipAddress2,ShipPhoneNumber,Title,TotalPrice,UsePoint,ShipCharge,TotalCouponDiscount,SettleAmount,Total_cal,BillMailAddress,TotalMallCouponDiscount,store_name,store_adress,store_zip,store_url,store_mail):
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
	subject = '【{store_name}】発送遅延のお知らせ'.format(store_name=store_name)
	message_text = """
----------------------------------------------------------------------
＝＝＝＝＝発送遅延が発生しております＝＝＝＝＝
----------------------------------------------------------------------
{BillLastName} {BillFirstName} 様

 {store_name}：担当の村上と申します。

この度はご注文頂きまして誠にありがとうございます。

首題の件に関しまして、誠に恐縮ではありますが
急激な注文数増加のため提携倉庫の処理が追い付かず発送処理に遅延が生じております。

本日中には発送できる見込みですので、暫しお時間頂戴すること御容赦お願い致します。

------------------------
＝＝＝＝＝御注文内容＝＝＝＝＝
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

◆◇◆　ポスト投函配送についてご確認ください　◆◇◆
ネコポス（ポスト投函）につきましては、配送方法をご選択いただきましても、同梱の組み合わせが対応サイズを超える場合は宅配便にてお届けとなります。
あらかじめご了承ください。
------------------------
------------------------

------------------------
  商品のお届けについて
------------------------
当日11時までの御注文の場合：当日発送
当日11時以降の御注文の場合：翌日発送
------------------------

------------------------

商品の発送が完了致しましたら、運送会社・伝票番号等をメールにてご連絡致します。
恐れ入りますが、次のご案内まで今しばらくお待ち下さい。

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
			TotalMallCouponDiscount=TotalMallCouponDiscount) 
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

def refresh_token(refresh_token):
	
	app_id = os.getenv("app_id")
	secret = os.getenv("secret")
	
	headers = {"Accept":"*/*","Content-Type":"application/x-www-form-urlencoded; charset=utf-8"}
	data = "grant_type=refresh_token&client_id={app_id}&client_secret={secret}&refresh_token={refresh_token}".format(app_id=app_id,secret=secret,refresh_token=refresh_token)
	
	url = "https://auth.login.yahoo.co.jp/yconnect/v2/token"

	req = requests.post(url=url,headers=headers,data=data)
	token_directry = json.loads(req.text)
	return token_directry["access_token"]

def yahoo_api_Order_acceptance(book_key,sh_list,refresh_token_base):
	
	#gspread側
	work_book = gspread_me(book_key)
	template_csv_path = "./入力/csv雛形/ヤマトFF_ShippingRequest.csv"
	template_csvs = []
	template_csvs = read_csv(template_csv_path)
	now = datetime.datetime.now()
	filename = './出力_csv/' + "yahoo_ヤマト発送依頼" + now.strftime('%Y%m%d_%H%M%S') + '.csv'
	header = [template_csvs[0]]
	write_csv(filename,header)
	
	Listing_sheet = work_book.worksheet(sh_list[1])
	Listing_all = Listing_sheet.get_all_values()
	
	directry = {}
	directry_size = {}
	Directory_Listing_Code = {}
	Directory_Listing_size = {}
	Directory_Listing_Title = {}
	for L_row in Listing_all:
		if L_row[18] == "TRUE":	
			directry[L_row[5]] =  L_row[14]
			directry_size[L_row[5]] =  L_row[19]
			Directory_Listing_Code[L_row[0]] = L_row[14]##
			Directory_Listing_size[L_row[0]] = L_row[19]##
			Directory_Listing_Title[L_row[0]] = L_row[6]##
			
	New_Order_sheet = work_book.worksheet(sh_list[6])
	New_Order_all = New_Order_sheet.get_all_values()
	Maxrow = len(New_Order_all)
	count = Maxrow + 300
	key = "A1:B" + str(count)
	order_range = New_Order_sheet.range(key)
	
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
		<Field>OrderId,ShipMethodName,OrderStatus,PayStatus,ItemId,ShipMethodName</Field>
	</Search>
	<SellerId>chan-mu</SellerId>
	</Req>""".format(OrderTimeFrom=T_from,OrderTimeTo=T_to)
		
	req = requests.post(url=url,headers=headers,data=data,cert=(#crtファイルを指定,#keyファイルを指定))	
	soup = BeautifulSoup(req.text, "xml")
	
	if not req.status_code == 200:
		sys.exit("リフレッシュトークンを更新して下さい")
		
	trees = []
	trees = soup.find_all("OrderInfo")
		
	OrderId_list = []
	for tree in trees:
		if tree.find("ShipMethodName").text == #処理対象の配送登録名を入力:
			if tree.find("OrderStatus").text == "2":
				if tree.find("PayStatus").text == "1":
					OrderId_list.append(tree.find("OrderId").text)
		
	url = "https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/orderInfo"
	headers = {"Authorization":"Bearer " + access_token}
	
	for OrderId in OrderId_list:
		data = """
		<Req>
			<Target>
			<OrderId>{OrderId}</OrderId>
			<Field>OrderId,ShipMethodName,OrderStatus,PayMethodName,PayStatus,Discount,BillMailAddress,BillFirstName,BillLastName,ItemId,UnitPrice,SubCode,OrderTime,Quantity,UnitPrice,TotalPrice,ShipFirstName,ShipLastName,ShipMethod,ShipZipCode,ShipPrefecture,ShipCity,ShipAddress1,ShipAddress2,ShipPhoneNumber,ShipRequestDate,ShipRequestTime,Title,Discount,UsePoint,TotalPrice,TotalCouponDiscount,SettleAmount,ShipCharge,TotalMallCouponDiscount</Field>
			</Target>
			<SellerId>chan-mu</SellerId>
		</Req>""".format(OrderId = OrderId)
		
		req = requests.post(url=url,headers=headers,data=data,(#crtファイルを指定,#keyファイルを指定))
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
		
		#新規注文シートにある場合、発送遅延
		delay_flag = False
		for New_Order_cell in New_Order_all:
			if New_Order_cell[0] == OrderId:
				print("新規注文シートに下記の注文IDが登録済みです")
				print("注文ID:" + str(OrderId))
				delay_flag = True
		
		if delay_flag:
			send_gmail_delay(BillFirstName,BillLastName,OrderId,OrderTime,PayMethodName,ShipMethodName,ShipRequestDate,ShipRequestTime,ShipLastName,ShipFirstName,ShipZipCode,ShipPrefecture,ShipCity,ShipAddress1,ShipAddress2,ShipPhoneNumber,Title,TotalPrice,UsePoint,ShipCharge,TotalCouponDiscount,SettleAmount,Total_cal,BillMailAddress,TotalMallCouponDiscount,store_name,store_adress,store_zip,store_url,store_mail)
			print("Delay:",OrderId)
			continue
		
		send_gmail_first(BillFirstName,BillLastName,OrderId,OrderTime,PayMethodName,ShipMethodName,ShipRequestDate,ShipRequestTime,ShipLastName,ShipFirstName,ShipZipCode,ShipPrefecture,ShipCity,ShipAddress1,ShipAddress2,ShipPhoneNumber,Title,TotalPrice,UsePoint,ShipCharge,TotalCouponDiscount,SettleAmount,Total_cal,BillMailAddress,TotalMallCouponDiscount,store_name,store_adress,store_zip,store_url,store_mail)
		print("SENT:",OrderId)
		
		#ヤマトcsv
		#時間もついてくるので成形
		OrderTime = datetime.datetime.strptime(OrderTime,'%Y-%m-%d %H:%M')
		OrderTime = OrderTime.strftime("%Y%m%d")
		ShipName = BillLastName + " " + BillFirstName
		#県・市町村が割れてるので合体
		shipAddress1 = ShipPrefecture + ShipCity + ShipAddress1
		Total = TotalPrice
		Total_fee = round(int(Total)*0.1)
			
		#出荷用csv成形
		##分割
		###return_sizeは"0101"を返すようにしてます。将来的にネコポスで可なのはネコポスを選択するように⇒対応済み
		###改定するために作成した関数です。
		for n,Item in enumerate(Items):
		
			ItemId_now = Item.find("ItemId").text
			QuantityDetail_now = Item.find("Quantity").text
			YamatoID_list = []
			if ItemId_now[:3] == "set" or ItemId_now[:3] == "SET":
				YamatoID_list = Change_SetID_to_YamatoID(work_book,sh_list,Directory_Listing_Code,Directory_Listing_size,Directory_Listing_Title,ItemId_now,QuantityDetail_now)
			else:
				try:
					YamatoID_list = [[directry[ItemId_now],QuantityDetail_now,directry_size[ItemId_now],Item.find("Title").text]]
				except:
					print("{}:continue".format(ItemId_now))
					continue
			
			UnitPrice_now = Item.find("UnitPrice").text
			LineSubTotal_now = int(UnitPrice_now) * int(QuantityDetail_now)
			Yamato_csv = []
			#csvの成形
			
			for data in YamatoID_list:
				Yamato_csv = [
				OrderId,
				return_size(data[2],ShipRequestDate),
				ShipRequestDate.replace("-",""),
				ShipRequestTime,
				ShipPhoneNumber,
				ShipZipCode.replace("-",""),
				shipAddress1,
				ShipAddress2,
				"",
				ShipName,
				"様",
				#ストア電話番号を記入,
				#ストア郵便番号を記入,
				#ストア住所を記入(市町村まで),
				#ストア住所を記入(市町村以降),
				"",
				#ストア名を記入,
				"",
				"雑貨",
				"",
				OrderTime,
				data[0],
				data[3],
				data[1],
				UnitPrice_now,
				LineSubTotal_now,
				Total,
				Total_fee,
				Total,
				"0",
				"0",
				UsePoint,
				Discount,
				"1",
				"",
				"",
				"",
				"",
				"",
				"",
				#ストア名を記入,
				#ストア郵便番号を記入,
				#ストア電話番号を記入,
				#ストア住所を記入,
				OrderId,
				ItemId_now,
				#ストア名を記入,
				"",
				"",
				Total,
				"",
				"",
				"",
				"1"]
				dummy = []
				dummy = [Yamato_csv]
				write_csv(filename,dummy)
		
		order_range[Maxrow*2].value = OrderId
		order_range[Maxrow*2+1].value = "yahoo"
		Maxrow += 1
	
	#Yamato_csvの配送サイズ指定を修正
	created_yamato_csv = []
	created_yamato_csv = read_csv(filename)
	os.remove(filename)
	count = 0
	
	code = {}
	before = None
	
	for i,cell in enumerate(created_yamato_csv):
		if i == 0:
			continue
		
		if before == cell[0]:
			code[cell[0]].append(cell[1])
		else:
			code[cell[0]] = []
			code[cell[0]].append(cell[1])
		
		before = cell[0]
	
	for i,cell in enumerate(created_yamato_csv):
		if i > 0:
			for n in code[cell[0]]:
				if n == "0101":
					cell[1] = "0101"
		
		write_csv(filename,[cell])
	
	if not i == 0:
		pathlib_filename = pathlib.Path(filename)
		upload_file = str(pathlib_filename.resolve())
		print(upload_file)
		loop = True
		
		while loop:
			loop = Yamato_shippinginfo_upload(upload_file)
	
	shutil.move(filename,"./出力_csv/処理済")	
	New_Order_sheet.update_cells(order_range)
	print("yahoo_api_Order_acceptance_Finish")

def au_api_Order_acceptance(book_key,sh_list,au_api_key):
	#gspread側
	work_book = gspread_me(book_key)
	template_csv_path = "./入力/csv雛形/ヤマトFF_ShippingRequest.csv"
	template_csvs = []
	template_csvs = read_csv(template_csv_path)
	now = datetime.datetime.now()
	filename = './出力_csv/' + "au_ヤマト発送依頼" + now.strftime('%Y%m%d_%H%M%S') + '.csv'
	header = [template_csvs[0]]
	write_csv(filename,header)
	Listing_sheet = work_book.worksheet(sh_list[1])
	Listing_all = Listing_sheet.get_all_values()
	
	directry = {}
	directry_size = {}
	Directory_Listing_Code = {}
	Directory_Listing_size = {}
	Directory_Listing_Title = {}
	for L_row in Listing_all:
		if L_row[18] == "TRUE":	
			directry[L_row[5]] =  L_row[14]
			directry_size[L_row[5]] =  L_row[19]
			Directory_Listing_Code[L_row[0]] = L_row[14]##
			Directory_Listing_size[L_row[0]] = L_row[19]##
			Directory_Listing_Title[L_row[0]] = L_row[6]##	
	
	New_Order_sheet = work_book.worksheet(sh_list[6])
	New_Order_all = New_Order_sheet.get_all_values()
	Maxrow = len(New_Order_all)
	count = Maxrow + 300
	key = "A1:B" + str(count)
	order_range = New_Order_sheet.range(key)
	
	#api側
	au_url = "https://api.manager.wowma.jp/wmshopapi/"
	T_now = datetime.datetime.now()
	T_before = T_now + datetime.timedelta(days=-30)
	
	endDate = T_now.strftime('%Y%m%d')
	startDate = T_before.strftime('%Y%m%d')
	requestURL = "https://api.manager.wowma.jp/wmshopapi/searchTradeInfoListProc?shopId=#ストアID(数字)&totalCount=100&startDate={startDate}&endDate={endDate}&orderStatus=新規受付"\
	.format(startDate=startDate,endDate=endDate)
	
	headers = {"Content-type": "application/x-www-form-urlencoded","Authorization": "Bearer " + au_api_key}
	req = requests.get(url=requestURL,headers=headers)
	soup = BeautifulSoup(req.text,"xml")
	
	orders = soup.find_all("orderInfo")
	for order in orders:
		if not order.find("deliveryName").text == #処理対象の配送登録名を入力:
			continue
		
		if not order.find("authorizationStatus").text == "Y":
			continue		
		
		#複数種注文の場合はdetailが複数ある
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
		
		#メール用とcsv用で2回for回すことになった。。。
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
		
		#新規注文シートにある場合は、発送遅延
		delay_flag = False
		for New_Order_cell in New_Order_all:
			if New_Order_cell[0] == OrderId:
				print("新規注文シートに下記の注文IDが登録済みです")
				print("注文ID:" + str(OrderId))
				delay_flag = True
		
		if delay_flag:
			send_gmail_delay(BillFirstName,BillLastName,OrderId,OrderTime,PayMethodName,ShipMethodName,ShipRequestDate,ShipRequestTime,ShipLastName,ShipFirstName,ShipZipCode,ShipPrefecture,ShipCity,ShipAddress1,ShipAddress2,ShipPhoneNumber,Title,TotalPrice,UsePoint,ShipCharge,TotalCouponDiscount,SettleAmount,Total_cal,BillMailAddress,TotalMallCouponDiscount,store_name,store_adress,store_zip,store_url,store_mail)
			print("Delay:",OrderId)
			continue
		
		send_gmail_first(BillFirstName,BillLastName,OrderId,OrderTime,PayMethodName,ShipMethodName,ShipRequestDate,ShipRequestTime,ShipLastName,ShipFirstName,ShipZipCode,ShipPrefecture,ShipCity,ShipAddress1,ShipAddress2,ShipPhoneNumber,Title,TotalPrice,UsePoint,ShipCharge,TotalCouponDiscount,SettleAmount,Total_cal,BillMailAddress,TotalMallCouponDiscount,store_name,store_adress,store_zip,store_url,store_mail)
		print("SENT:",OrderId)
		
		#出荷用csv成形
		##分割
		
		for detail in details:
			ItemId_now = detail.find("itemCode").text
			QuantityDetail_now = detail.find("unit").text
			
			YamatoID_list = []
			if ItemId_now[:3] == "set" or ItemId_now[:3] == "SET":
				YamatoID_list = Change_SetID_to_YamatoID(work_book,sh_list,Directory_Listing_Code,Directory_Listing_size,Directory_Listing_Title,ItemId_now,QuantityDetail_now)
			else:
				YamatoID_list = [[directry[ItemId_now],QuantityDetail_now,directry_size[ItemId_now],detail.find("itemName").text]]

			UnitPrice_now = detail.find("itemPrice").text
			LineSubTotal_now = detail.find("totalItemPrice").text
			
			Yamato_csv = []
			#csvの成形
			for data in YamatoID_list:
				Yamato_csv = [
				OrderId,
				return_size(data[2],ShipRequestDate),
				ShipRequestDate,
				ShipRequestTime,
				ShipPhoneNumber,
				ShipZipCode.replace("-",""),
				shipAddress1,
				shipAddress2,
				"",
				ShipName,
				"様",
				#ストア電話番号を記入,
				#ストア郵便番号を記入,
				#ストア住所を記入(市町村まで),
				#ストア住所を記入(市町村以降),
				"",
				#ストア名を記入,
				"",
				"雑貨",
				"",
				OrderTime,
				data[0],
				data[3],
				data[1],
				UnitPrice_now,
				LineSubTotal_now,
				Total,
				Total_fee,
				Total,
				"0",
				"0",
				UsePoint,
				Discount,
				"1",
				"",
				"",
				"",
				"",
				"",
				"",
				#ストア名を記入,
				#ストア郵便番号を記入,
				#ストア電話番号を記入,
				#ストア住所を記入,
				OrderId,
				ItemId_now,
				#ストア名を記入,
				"",
				"",
				Total,
				"",
				"",
				"",
				"4"]
				dummy = []
				dummy = [Yamato_csv]
				write_csv(filename,dummy)
		
		if not order_range[Maxrow*2-2].value == OrderId:
			order_range[Maxrow*2].value = OrderId
			order_range[Maxrow*2+1].value = "au"
			Maxrow += 1


	#Yamato_csvの配送サイズ指定を修正
	created_yamato_csv = []
	created_yamato_csv = read_csv(filename)
	os.remove(filename)
	count = 0
	
	code = {}
	before = None
	
	for i,cell in enumerate(created_yamato_csv):
		if i == 0:
			continue
		
		if before == cell[0]:
			code[cell[0]].append(cell[1])
		else:
			code[cell[0]] = []
			code[cell[0]].append(cell[1])
		
		before = cell[0]
	
	for i,cell in enumerate(created_yamato_csv):
		if i > 0:
			for n in code[cell[0]]:
				if n == "0101":
					cell[1] = "0101"
		
		write_csv(filename,[cell])
	
	if not i == 0:
		pathlib_filename = pathlib.Path(filename)
		upload_file = str(pathlib_filename.resolve())
		print(upload_file)
		loop = True
		
		while loop:
			loop = Yamato_shippinginfo_upload(upload_file)
	
	shutil.move(filename,"./出力_csv/処理済")	
	New_Order_sheet.update_cells(order_range)
	print("Creat_au_shipping_csv_Finish")

def Qoo10_Order_Acceptance(book_key,sh_list,Qoo10_SAK):
	#gspread側
	work_book = gspread_me(book_key)
	template_csv_path = "./入力/csv雛形/ヤマトFF_ShippingRequest.csv"
	template_csvs = []
	template_csvs = read_csv(template_csv_path)
	now = datetime.datetime.now()
	filename = './出力_csv/' + "Qoo10_ヤマト発送依頼" + now.strftime('%Y%m%d_%H%M%S') + '.csv'
	header = [template_csvs[0]]
	write_csv(filename,header)
	
	Listing_sheet = work_book.worksheet(sh_list[1])
	Listing_all = Listing_sheet.get_all_values()
	
	directry = {}
	directry_size = {}
	Directory_Listing_Code = {}
	Directory_Listing_size = {}
	Directory_Listing_Title = {}
	for L_row in Listing_all:
		if L_row[18] == "TRUE":	
			directry[L_row[5]] =  L_row[14]
			directry_size[L_row[5]] =  L_row[19]
			Directory_Listing_Code[L_row[0]] = L_row[14]##
			Directory_Listing_size[L_row[0]] = L_row[19]##
			Directory_Listing_Title[L_row[0]] = L_row[6]##
			
	New_Order_sheet = work_book.worksheet(sh_list[6])
	New_Order_all = New_Order_sheet.get_all_values()
	Maxrow = len(New_Order_all)
	count = Maxrow + 300
	key = "A1:B" + str(count)
	order_range = New_Order_sheet.range(key)

	#api側
	T_now = datetime.datetime.now()
	T_before = T_now + datetime.timedelta(days=-30)
	
	endDate = T_now.strftime('%Y%m%d')
	startDate = T_before.strftime('%Y%m%d')
	requestURL = "https://api.qoo10.jp/GMKT.INC.Front.QAPIService/ebayjapan.qapi??v=1.0&returnType=xml&&ShippingStat=2&method=ShippingBasic.GetShippingInfo_v2&key={Qoo10_SAK}&search_Sdate={startDate}&search_Edate={endDate}"\
	.format(Qoo10_SAK=Qoo10_SAK,startDate=startDate,endDate=endDate)
	req = requests.get(url=requestURL)
	soup = BeautifulSoup(req.text,"xml")
	orders = soup.find_all("ShippingInfo")
	data = {}
	print(orders)
	for order in orders:
		
		shippingStatus = order.find("shippingStatus").text
		OrderId = order.find("packNo").text
				
		if not OrderId in data:
			data[OrderId] = {}
			data[OrderId]["Title"] = ""
			data[OrderId]["Total_cal"] = 0
		
		
		if not order.find("DeliveryCompany").text == #処理対象の配送登録名を入力:
			continue
		
				
			
		data[OrderId]["packNo"] = order.find("packNo").text
		data[OrderId]["BillFirstName"] = order.find("buyer").text
		OrderTime = order.find("orderDate").text
		OrderTime = datetime.datetime.fromisoformat(OrderTime)
		data[OrderId]["OrderTime"] = datetime.datetime.strftime(OrderTime,'%Y-%m-%d %H:%M:%S')
		data[OrderId]["BillMailAddress"] = order.find("buyerEmail").text
		data[OrderId]["PayMethodName"] = order.find("PaymentMethod").text
		data[OrderId]["ShipFirstName"] = order.find("receiver").text
		data[OrderId]["ShipMethodName"] = "ヤマト運輸"
		data[OrderId]["ShipZipCode"] = order.find("zipCode").text.replace("-","")
		data[OrderId]["ShipAddress1"] = order.find("Addr1").text
		data[OrderId]["ShipAddress2"] = order.find("Addr2").text
		data[OrderId]["ShipPhoneNumber"] = order.find("receiverMobile").text.replace("+81","").replace("-","")
		data[OrderId]["Discount"] = order.find("Cart_Discount_Seller").text.replace(".0000","")
		data[OrderId]["TotalMallCouponDiscount"] = int(order.find("Cart_Discount_Qoo10").text.replace(".0000",""))
		data[OrderId]["Title"] += order.find("itemTitle").text + "　:　" + order.find("orderPrice").text.replace(".0000","") + "円　：　" + order.find("orderQty").text + "個" + "\n"
		data[OrderId]["Total_cal"] += int(order.find("total").text.replace(".0000",""))
			
		#スプレッドシートに書き込み
		order_range[Maxrow*2].value = OrderId + "," + order.find("orderNo").text
		order_range[Maxrow*2+1].value = "Qoo10"
		Maxrow += 1
			
		print(order)
		#YFF_csv作成
		ItemId_now = order.find("sellerItemCode").text
		QuantityDetail_now = order.find("orderQty").text
		YamatoID_list = []
		if ItemId_now[:3] == "set" or ItemId_now[:3] == "SET":
			YamatoID_list = Change_SetID_to_YamatoID(work_book,sh_list,Directory_Listing_Code,Directory_Listing_size,Directory_Listing_Title,ItemId_now,QuantityDetail_now)
		else:
			YamatoID_list = [[directry[ItemId_now],QuantityDetail_now,directry_size[ItemId_now],order.find("itemTitle").text]]
			
		UnitPrice_now = order.find("orderPrice").text.replace(".0000","")
		LineSubTotal_now = int(UnitPrice_now) * int(QuantityDetail_now)
		Yamato_csv = []
		#csvの成形
		for row in YamatoID_list:
			Yamato_csv = [
			"Q" + OrderId,
			return_size(row[2],""),
			"",
			"",
			data[OrderId]["ShipPhoneNumber"],
			data[OrderId]["ShipZipCode"],
			data[OrderId]["ShipAddress1"],
			data[OrderId]["ShipAddress2"],
			"",
			data[OrderId]["ShipFirstName"],
			"様",
			#ストア電話番号を記入,
			#ストア郵便番号を記入,
			#ストア住所を記入(市町村まで),
			#ストア住所を記入(市町村以降),
			"",
			#ストア名を記入,
			"",
			"雑貨",
			"",
			datetime.datetime.strftime(OrderTime,'%Y%m%d'),
			row[0],
			row[3],
			row[1],
			UnitPrice_now,
			LineSubTotal_now,
			"0",
			"0",
			"0",
			"0",
			"0",
			"0",
			"0",
			"1",
			"",
			"",
			"",
			"",
			"",
			"",
			#ストア名を記入,
			#ストア郵便番号を記入,
			#ストア電話番号を記入,
			#ストア住所を記入,
			OrderId,
			ItemId_now,
			#ストア名を記入,
			"",
			"",
			"",
			"",
			"",
			"",
			"4"]
			write_csv(filename,[Yamato_csv])
				
				
		store_name = #ストア名を記入
		store_adress = #ストア住所を記入
		store_zip = #ストア郵便番号を記入
		store_url = #ストアURLを記入
		store_mail = #メールアドレスを記入
	
	for cart in data.values():
		
		print(cart)
		delay_flag = False
		for New_Order_cell in New_Order_all:
			if New_Order_cell[0] == OrderId:
				print("新規注文シートに下記の注文IDが登録済みです")
				print("注文ID:" + str(OrderId))
				delay_flag = True
		
		BillLastName = ""
		ShipRequestDate = ""
		ShipRequestTime = ""
		ShipLastName = ""
		ShipPrefecture = ""
		ShipCity = ""
		TotalPrice = ""
		UsePoint = "注文履歴より確認お願い致します-"
		ShipCharge = "0"
		SettleAmount = ""
	
		if delay_flag:
			send_gmail_delay(
			cart["BillFirstName"],
			BillLastName,
			cart["packNo"] + "(カート番号)",
			cart["OrderTime"],
			cart["PayMethodName"],
			cart["ShipMethodName"],
			ShipRequestDate,
			ShipRequestTime,
			ShipLastName,
			cart["ShipFirstName"],
			cart["ShipZipCode"],
			ShipPrefecture,
			ShipCity,
			cart["ShipAddress1"],
			cart["ShipAddress2"],
			cart["ShipPhoneNumber"],
			cart["Title"],
			TotalPrice,
			UsePoint,
			ShipCharge,
			cart["TotalMallCouponDiscount"],
			cart["Total_cal"] - cart["TotalMallCouponDiscount"],
			cart["Total_cal"],
			cart["BillMailAddress"],
			cart["TotalMallCouponDiscount"],
			store_name,
			store_adress,
			store_zip,
			store_url,
			store_mail)
			print("Delay:",OrderId)
			continue
		
		send_gmail_first(
		cart["BillFirstName"],
		BillLastName,
		cart["packNo"] + "(カート番号)",
		cart["OrderTime"],
		cart["PayMethodName"],
		cart["ShipMethodName"],
		ShipRequestDate,
		ShipRequestTime,
		ShipLastName,
		cart["ShipFirstName"],
		cart["ShipZipCode"],
		ShipPrefecture,
		ShipCity,
		cart["ShipAddress1"],
		cart["ShipAddress2"],
		cart["ShipPhoneNumber"],
		cart["Title"],
		TotalPrice,
		UsePoint,
		ShipCharge,
		cart["TotalMallCouponDiscount"],
		cart["Total_cal"] - cart["TotalMallCouponDiscount"],
		cart["Total_cal"],
		cart["BillMailAddress"],
		cart["TotalMallCouponDiscount"],
		store_name,
		store_adress,
		store_zip,
		store_url,
		store_mail)
			
		print("SENT:",OrderId)
		
	#Yamato_csvの配送サイズ指定を修正
	created_yamato_csv = []
	created_yamato_csv = read_csv(filename)
	os.remove(filename)
	count = 0
	
	code = {}
	before = None
	
	for i,cell in enumerate(created_yamato_csv):
		if i == 0:
			continue
		
		if before == cell[0]:
			code[cell[0]].append(cell[1])
		else:
			code[cell[0]] = []
			code[cell[0]].append(cell[1])
		
		before = cell[0]
	
	for i,cell in enumerate(created_yamato_csv):
		if i > 0:
			for n in code[cell[0]]:
				if n == "0101":
					cell[1] = "0101"
		
		write_csv(filename,[cell])
	
	if not i == 0:
		pathlib_filename = pathlib.Path(filename)
		upload_file = str(pathlib_filename.resolve())
		print(upload_file)
		loop = True
		
		while loop:
			loop = Yamato_shippinginfo_upload(upload_file)
	
	shutil.move(filename,"./出力_csv/処理済")	
	New_Order_sheet.update_cells(order_range)
	print("Qoo10_Order_Acceptance_Finish")

au_api_key = os.getenv("au_api_key")
refresh_token_base = os.getenv("refresh_token_base")
Qoo10_SAK = os.getenv("Qoo10_SAK")
#ここにリフレッシュトークン入れる

def Master_api_Order_acceptance(book_key,sh_list,refresh_token_base,au_api_key,Qoo10_SAK):
	yahoo_api_Order_acceptance(book_key,sh_list,refresh_token_base)
	au_api_Order_acceptance(book_key,sh_list,au_api_key)
	Qoo10_Order_Acceptance(book_key,sh_list,Qoo10_SAK)

book_key = os.getenv("CosmeProduct")
sh_list = ["仕入れ商品リスト","出品商品リスト","仕入","引当","在庫","受注","新規注文","セット商品リスト"]

Master_api_Order_acceptance(book_key,sh_list,refresh_token_base,au_api_key,Qoo10_SAK)