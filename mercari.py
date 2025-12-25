from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import re
import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# 変数定義
sheet_headers = ["価格", "コメント", "画像", "URL"]
target_url = 'https://jp.mercari.com/'
sheet_title = "mercari_scraping_results"

# スプレッドシート設定
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
# credentials.json はご自身の認証ファイル名に合わせてください
creds = ServiceAccountCredentials.from_json_keyfile_name('./resale-482213-619ca3d89121.json', scope)
client = gspread.authorize(creds)

# 既存のスプレッドシートを開く (ファイル名を指定)
# 事前にGoogleドライブで「mercari_scraping_results」という名前のシートを作成し、
# サービスアカウント(yna-755@...)を編集者として招待しておく必要があります。
try:
    spreadsheet = client.open(sheet_title)
    sheet = spreadsheet.sheet1
    print("既存のスプレッドシートを開きました")
    
    # シートの中身を全消去
    sheet.clear()
    print("シートの内容をクリアしました")
    
    # ヘッダーを追加
    sheet.append_row(sheet_headers)
except Exception as e:
    print(f"スプレッドシートを開けませんでした: {e}")
    exit()

# Chrome のオプションを設定する
options = webdriver.ChromeOptions()
# options.add_argument('--headless') # 画面を見るためにコメントアウト
options.add_experimental_option("detach", True) # ブラウザを自動で閉じない設定

# ローカルの Chrome を使用する
driver = webdriver.Chrome(options=options)

# Selenium 経由でブラウザを操作する
driver.get(target_url)

# モーダルが表示されるのを待つ
time.sleep(3)

# モーダルを閉じるために ESC キーを送信する
try:
    webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
    print("ESCキーを送信しました")
except Exception as e:
    print(f"ESCキー送信失敗: {e}")

# 要素を直接取得する
try:
    # クラス名で検索ボタンを指定する
    search_button = driver.find_element(By.CSS_SELECTOR, '.iconButton__d01ad679.navigation__d01ad679.circle__d01ad679')

    search_button.click()
    print("検索ボタンをクリックしました")

    # 同じクラスを持つ要素をすべて取得する (find_elements と複数形にする)
    inputs = driver.find_elements(By.CSS_SELECTOR, ".sc-7209bb23-2.kpWntS")
    
    found_visible = False
    for search_input in inputs:
        # 表示されている要素かチェックする
        if search_input.is_displayed():
            print("表示されている入力欄が見つかりました")
            search_input.clear()
            search_input.send_keys("apple pencil pro")
            found_visible = True
            break # 見つかったらループを抜ける
            
    if not found_visible:
        print("表示されている入力欄が見つかりませんでした")

    seach_icon_in_search_text_area = driver.find_element(By.CSS_SELECTOR, '.iconButton__d01ad679.transparent__d01ad679[aria-label="検索"]')
    seach_icon_in_search_text_area.click()
    print("検索アイコンをクリックしました")

    # 検索結果が表示されるまで待つ
    time.sleep(5)

    # 一覧取得
    thumbnails = driver.find_elements(By.CSS_SELECTOR, '.merItemThumbnail')
    print(f"取得した商品数: {len(thumbnails)}")

    # 書き込み用リスト
    rows_to_add = []

    # 上記要素をループで回し、id属性がmerItemThumbnail を抽出する
    for thumbnail in thumbnails:
        try:
            # 上記要素のaria-labelを取得
            aria_label = thumbnail.get_attribute("aria-label")
            
            if aria_label:
                # 正規表現で「数字+円」のパターンを探す（位置に関係なく抽出）
                match = re.search(r'(\d{1,6}(?:,\d{6})*)円', aria_label)
                
                if match:
                    # カンマを除去して数値文字列にする
                    price_str = match.group(1).replace(",", "")
                else:
                    price_str = "0"
                
            #    閾値の金額以下の場合、スプレッドシートに出力
                
                # 数値に変換
                price = int(price_str)

                # thumbnail要素の親要素（リンク）を取得する
                product_url_elm = thumbnail.find_element(By.XPATH, './ancestor::a[@data-testid="thumbnail-link"]')
                product_url = product_url_elm.get_attribute("href")

                # コメント
                
                
                # 閾値（例: 20000円）以下の場合
                if price <= 20000:
                    print(f"閾値以下の商品を発見: {price}円 - {aria_label}")
                    # リストに追加するだけ（APIは呼ばない）
                    rows_to_add.append([aria_label, price, product_url])
                
        except Exception as inner_e:
            # 個別の要素取得失敗はスキップして続行
            continue

    # ループが終わった後に、まとめて書き込む
    if rows_to_add:
        sheet.append_rows(rows_to_add)
        print(f"{len(rows_to_add)}件のデータをスプレッドシートに書き込みました")
    else:
        print("条件に合う商品は見つかりませんでした")

except Exception as e:
    print(f"エラーが発生しました: {e}")
