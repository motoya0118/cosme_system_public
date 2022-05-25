# 化粧品販売システム

## ソースコード公開の目的
**自身のポートフォリオとして公開するため**

※MITライセンスで公開しておりますので、本ソースコードはご自由に御使い下さい

## ソースコード閲覧頂く方へ

*閲覧頂き大変感謝申し上げます*

**各ファイル内の主要な関数を下記に挙げますので、こちらのみ確認頂ければ主要な機能が確認可能です。**

### AM11_order.py
**yahoo_api_Order_acceptance(book_key,sh_list,refresh_token_base)**
**au_api_Order_acceptance(book_key,sh_list,au_api_key)**
**Qoo10_Order_Acceptance(book_key,sh_list,Qoo10_SAK)**

⇒注文情報を取得して、倉庫に発送指示・顧客へ注文承諾通知メール送信する関数です。

※発送が遅れてる場合は、発送遅延通知メールを送信します。

### PM1_UpdateStock.py
**Update_Stock(book_key,sh_list,refresh_token_base,au_api_key)**

⇒倉庫から在庫情報を取得して、各モールに在庫を反映する関数です。

### Moniter.py
**Inventory_Monitoring(book_key,sh_list,refresh_token_base,au_api_key)**

⇒注文情報を照会して、全モール合算で在庫0になった商品は全モール在庫0にする関数です。

### PM7_Master_shipping.py
**Master_Shipping_Notification(book_key,sh_list,au_api_key,refresh_token_base,Qoo10_SAK)**

⇒発送通知を出す関数です。

**puchage_allocation(book_key,sh_list)**

⇒データベース(仕入 → 引当)内処理で、仕入れ情報を在庫管理しやすい形に変換して引当テーブルに転記する関数です。

**order_organize_yahoo(book_key,sh_list,refresh_token_base)**
**order_organize_au(book_key,sh_list,au_api_key)**
**order_organize_Qoo10(book_key,sh_list,Qoo10_SAK)**

⇒注文情報を照会して、発送完了した注文をデータベース(受注)へ転記する関数です。

**Order_Allocation(book_key,sh_list)**

⇒データベース(受注 → 引当)内処理で、注文商品情報を在庫管理しやすい形に変換して引当テーブルに転記する関数です。

**stock_allocation(book_key,sh_list)**

⇒データベース(引当 → 在庫)内処理で、目視で在庫管理しやすい形に変換して在庫テーブルに出力する関数です。

**Master_after_Notification(book_key,sh_list,au_api_key,refresh_token_base)**

⇒発送から3日以上経過した注文について、レビュー依頼メールを送信する関数です。
