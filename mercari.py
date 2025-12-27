# todo：
# ・スクロールを回数固定ではなく動的に
# ・スプシには動的に書き込みしてすぐに購入できるように

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
search_title = "apple pencil pro 純正"
scrolle_volume = "1000"
scroll_count = 15  # 10回スクロールする（必要に応じて増やしてください）
current_scroll_position = 0 # 現在のスクロール位置
scroll_step = 500 # 1回のスクロール量 (ピクセル)
price_min = 300
price_max = 10000

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
try:
    driver = webdriver.Chrome(options=options)
except Exception as e:
    print(f"failed webdriver: {e}")

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

    # 絞り込みボタン押下
    # 絞り込みボタンをクリック
    try:
        filter_button = driver.find_element(By.CSS_SELECTOR, '.merButton.secondary__01a6ef84.small__01a6ef84')
        filter_button.click()
        print("絞り込みボタンをクリックしました")
    except Exception as e:
        print(f"絞り込みボタンのクリックに失敗しました: {e}")

    # 表示する商品の条件絞り込み
    # 条件タグと詳細のマッピング作ってそれをループでやる
    search_narrow_downs = [
        ["item_types", "mercari"], # 出品者
        ["price", "price"], # 価格
        ["d664efe3-ae5a-4824-b729-e789bf93aba9", "B38F1DC9286E0B80812D9B19DB14298C1FF1116CA8332D9EE9061026635C9088"], # 出品形式
    ]
    for search_narrow_down in search_narrow_downs:
        try:
            header_id = search_narrow_down[0]
            body_value = search_narrow_down[1]
            
            print(f"絞り込み処理開始: {header_id} -> {body_value}")

            # 1. 親要素(li)が表示されるまで待つ (最大10秒)
            li_elm = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, f"[data-testid='{header_id}']"))
            )

            print(li_elm)
            # 2. クリックすべきアコーディオンヘッダーを探す
            # .merAccordion が見つからない場合は li_elm 自体をターゲットにする
            try:
                header_clickable = li_elm.find_element(By.CSS_SELECTOR, ".merAccordion")
            except:
                header_clickable = li_elm

            # 4. ヘッダーをクリックして開く
            header_clickable.click()
            print(f"ヘッダーをクリックしました")
            time.sleep(1) # アニメーション待ち

            # 価格対応
            if header_id == "price":
                price_min_elm = li_elm.find_element(By.CSS_SELECTOR, f"input[name='priceMin']")
                price_min_elm.send_keys(price_min)
                price_max_elm = li_elm.find_element(By.CSS_SELECTOR, f"input[name='priceMax']")
                price_max_elm.send_keys(price_max)
                continue

            # 5. チェックボックスを選択
            # inputタグを探す
            body_input = li_elm.find_element(By.CSS_SELECTOR, f"input[value='{body_value}']")
            
            # JavaScriptで強制クリック (inputが隠れていても効く)
            driver.execute_script("arguments[0].click();", body_input)
            print(f"値 {body_value} を選択しました")
            
            # time.sleep(3) # 絞り込み反映待ち

        except Exception as e:
            print(f"検索絞り込みで失敗 ({search_narrow_down}): {e}")
            continue
    
    # 絞り込みメニューを閉じる（完了ボタンなどをクリック）
    try:
        close_filter_button = driver.find_element(By.CSS_SELECTOR, '.header__1d92fe3f > .merIconButton')
        close_filter_button.click()
        print("絞り込みメニューを閉じました")
        time.sleep(2) # メニューが閉じてリストが更新されるのを待つ
    except Exception as e:
        print(f"絞り込みメニューを閉じるボタンのクリックに失敗しました: {e}")

    # --- 修正: 動的スクロール (徐々にスクロール) ---
    print("スクロールを開始します...")

    # 変化がなかった回数をカウントして無限ループ防止
    no_change_count = 0
    max_no_change = 3  # 3回連続で高さが変わらなければ終了

    while True:
        # 現在のページ高さを取得
        last_height = driver.execute_script("return document.body.scrollHeight")

        # 少しずつスクロールする
        current_scroll_position += scroll_step
        driver.execute_script(f"window.scrollTo(0, {current_scroll_position});")
        
        # 読み込み待ち (少し短めでOK)
        time.sleep(0.5)

        # 現在の高さがスクロール位置より大きいか確認（まだ下にコンテンツがあるか）
        # または、スクロール位置がページ高さを超えた場合に、新しいコンテンツが読み込まれたか確認
        new_height = driver.execute_script("return document.body.scrollHeight")
        
        # スクロール位置がページ最下部付近に達したかチェック
        if current_scroll_position >= new_height:
            # 最下部に達したと思われる場合、少し待って本当に増えないか確認
            time.sleep(2)
            new_height_after_wait = driver.execute_script("return document.body.scrollHeight")
            
            if new_height_after_wait == last_height:
                no_change_count += 1
                print(f"最下部付近待機: {no_change_count}/{max_no_change}")
                if no_change_count >= max_no_change:
                    print("ページの最下部に到達しました")
                    break
            else:
                # 高さが伸びたのでリセットして続行
                no_change_count = 0
                print("新しいコンテンツが読み込まれました")
        
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

            # # メルカリショップは省く
            # if product_url and re.search(rf"{re.escape(target_url)}shops", product_url):
            #     continue

            product_urls.append(product_url)
        except Exception as inner_e:
            # 個別の要素取得失敗はスキップして続行
            continue

    print(f"取得した商品数: {len(product_urls)}")
    
    # 商品詳細画面で必要情報取得
    for product_url in product_urls:
        try:
            driver.get(product_url)
            print(product_url)
            time.sleep(2) # ページ遷移待ち

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

            # スプシに行を追加 (即時反映)
            sheet.append_row([price, comment, image_formula, product_url], value_input_option='USER_ENTERED')
            print(f"スプレッドシートに書き込みました: {price}")

        except Exception as inner_e:
            print(f"商品詳細の情報取得に失敗しました: {inner_e}")
            continue

except Exception as e:
    print(f"エラーが発生しました: {e}")
