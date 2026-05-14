# テーブル定義書

システム名：請求書管理システム  
RDBMS：PostgreSQL 18  
スキーマ：public  

---

## my_company（発行者情報）

| No. | 物理名 | データ型 | Not Null | デフォルト | 備考 |
|---|---|---|---|---|---|
| 1 | company_id | serial | Yes (PK) | | |
| 2 | company_name | varchar(100) | | | |
| 3 | postal_code | varchar(10) | | | |
| 4 | address1 | varchar(200) | | | 住所1 |
| 5 | address2 | varchar(200) | | | 住所2 |
| 6 | tel | varchar(20) | | | |
| 7 | fax | varchar(20) | | | |
| 8 | invoice_number | varchar(20) | | | インボイス登録番号 |
| 9 | staff_name | varchar(50) | | | |
| 10 | bank_name | varchar(50) | | | 振込先銀行名・支店名 |
| 11 | account_type | varchar(30) | | | 口座種別 |
| 12 | account_number | varchar(20) | | | 口座番号 |
| 13 | account_name | varchar(60) | | | 口座名義 |

---

## clients（取引先）

| No. | 物理名 | データ型 | Not Null | デフォルト | 備考 |
|---|---|---|---|---|---|
| 1 | client_id | serial | Yes (PK) | | |
| 2 | client_name | varchar(100) | | | |
| 3 | postal_code | varchar(10) | | | |
| 4 | address1 | varchar(200) | | | 住所1 |
| 5 | address2 | varchar(200) | | | 住所2 |
| 6 | tel | varchar(20) | | | |
| 7 | email | varchar(100) | | | |
| 8 | contact_name | varchar(50) | | | 取引先担当者名 |
| 9 | is_deleted | boolean | | false | 削除済みフラグ（論理削除） |

**外部キー**：orders.client_id → clients.client_id

---

## products（商品）

| No. | 物理名 | データ型 | Not Null | デフォルト | 備考 |
|---|---|---|---|---|---|
| 1 | prod_id | serial | Yes (PK) | | |
| 2 | product_name | varchar(100) | | | |
| 3 | product_category | varchar(50) | | | |
| 4 | product_subcategory | varchar(50) | | | 商品小分類 |
| 5 | unit | varchar(20) | | | |
| 6 | price | integer | | | |
| 7 | tax_rate | varchar(5) | | | デフォルト税率（10%または8%） |
| 8 | stock | integer | | | |
| 9 | reorder_point | integer | | | 在庫不足アラートの基準値 |

**外部キー**：order_details.prod_id → products.prod_id

---

## orders（注文）

| No. | 物理名 | データ型 | Not Null | デフォルト | 備考 |
|---|---|---|---|---|---|
| 1 | order_id | serial | Yes (PK) | | |
| 2 | client_id | integer | | | |
| 3 | order_number | varchar(50) | | | |
| 4 | order_date | date | | | |
| 5 | subject | varchar(200) | | | 件名 |
| 6 | delivery_date | date | | | 納期 |
| 7 | expiry_date | date | | | 見積書の有効期限 |
| 8 | payment_terms | varchar(100) | | | 支払条件 |
| 9 | payment_due_date | date | | | 支払期限 |
| 10 | is_quoted | boolean | | false | 見積書発行済みフラグ |
| 11 | is_delivered | boolean | | false | 納品書発行済みフラグ |
| 12 | is_invoiced | boolean | | false | 請求書発行済みフラグ |
| 13 | is_receipted | boolean | | false | 領収書発行済みフラグ |
| 14 | is_stock_synced | boolean | | false | 在庫同期済みフラグ |

**外部キー**：orders.client_id → clients.client_id

---

## order_details（注文明細）

| No. | 物理名 | データ型 | Not Null | デフォルト | 備考 |
|---|---|---|---|---|---|
| 1 | detail_id | serial | Yes (PK) | | |
| 2 | order_id | integer | | | |
| 3 | prod_id | integer | | | |
| 4 | quantity | integer | | | |
| 5 | unit_price | integer | | | |
| 6 | tax_rate | varchar | | | |
| 7 | free_item_name | varchar(100) | | | マスタ未登録商品名 |

**外部キー**：order_details.order_id → orders.order_id  
**外部キー**：order_details.prod_id → products.prod_id
