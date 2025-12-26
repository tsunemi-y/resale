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
search_title = "apple"
scrolle_volume = "1000"
scroll_count = 10  # 10回スクロールする（必要に応じて増やしてください）

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
            search_input.send_keys(search_title)
            found_visible = True
            break # 見つかったらループを抜ける
            
    if not found_visible:
        print("表示されている入力欄が見つかりませんでした")

    seach_icon_in_search_text_area = driver.find_element(By.CSS_SELECTOR, '.iconButton__d01ad679.transparent__d01ad679[aria-label="検索"]')
    seach_icon_in_search_text_area.click()
    print("検索アイコンをクリックしました")

    # 検索結果が表示されるまで待つ
    time.sleep(3)

    # --- 修正: 回数指定で確実にスクロールする ---
    print("スクロールを開始します...")
    
    for i in range(scroll_count):
        # 現在のスクロール位置から、少しずつ下に移動させる
        # 1000だと５回で最深部までいく
        driver.execute_script(f"window.scrollBy(0, {scrolle_volume});")
        time.sleep(1)
        driver.execute_script(f"window.scrollBy(0, {scrolle_volume});")
        time.sleep(1)
        
        # 最後に最下部へダメ押し
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        
        print(f"スクロール中... ({i+1}/{scroll_count})")
        time.sleep(4) # 読み込み待ち
    
    print("スクロール終了")
    # --- 修正終了 ---

    # 一覧取得
    thumbnails = driver.find_elements(By.CSS_SELECTOR, '.merItemThumbnail')

    # 上記要素をループで回し、URLリスト取得
    product_urls = []
    for thumbnail in thumbnails:
        try:
            # thumbnail要素の親要素（リンク）を取得する
            product_url_elm = thumbnail.find_element(By.XPATH, './ancestor::a[@data-testid="thumbnail-link"]')
            product_url = product_url_elm.get_attribute("href")

            # メルカリショップは省く
            if product_url and re.search(rf"{re.escape(target_url)}shops", product_url):
                continue

            product_urls.append(product_url)
        except Exception as inner_e:
            # 個別の要素取得失敗はスキップして続行
            continue

    print(f"取得した商品数: {len(product_urls)}")
    
    # 商品詳細画面で必要情報取得
    rows_to_add = []
    for product_url in product_urls:
        try:
            driver.get(product_url)
            print(product_url)
            time.sleep(2) # ページ遷移待ち
            # ["価格", "コメント", "画像", "URL"]

            # オークション商品除外
            tag_elm = driver.find_element(By.CSS_SELECTOR, ".tagText__5000b0c4")
            tag_text = tag_elm.text
            if tag_text == 'オークション商品':
                continue

            # 価格
            price_elm = driver.find_element(By.CSS_SELECTOR, "[data-testid='price'] > span:nth-of-type(2)")
            price = price_elm.text

            # コメント
            comment_elm = driver.find_element(By.CSS_SELECTOR, "[data-testid='description']")
            comment = comment_elm.text

            # 画像
            img_elm = driver.find_element(By.CSS_SELECTOR, "[data-testid='image-0']").find_element(By.TAG_NAME, "img")
            print(img_elm)
            src = img_elm.get_attribute("src")
            
            # URLをIMAGE関数で囲む
            image_formula = f'=IMAGE("{src}", 4, 200, 200)'

            # スプシの行作成 (src ではなく image_formula を入れる)
            rows_to_add.append([price, comment, image_formula, product_url])

        except Exception as inner_e:
            print(f"商品詳細の情報取得に失敗しました: {inner_e}")
            continue

    # スプレッドシートにまとめて書き込む
    if rows_to_add:
        # value_input_option='USER_ENTERED' を指定することで、=IMAGE() が数式として機能する
        sheet.append_rows(rows_to_add, value_input_option='USER_ENTERED')
        print(f"{len(rows_to_add)}件のデータをスプレッドシートに書き込みました")
    else:
        print("条件に合う商品は見つかりませんでした")

except Exception as e:
    print(f"エラーが発生しました: {e}")
