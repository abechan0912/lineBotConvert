import logging
import time
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

# 採用管理サイトクラスを定義
class RecruitWebSite:

    # コンストラクタを定義
    def __init__(self, downloadDirectory, driverPath, accessUrl, loginId, loginPass) -> None:
        # ドライバー設定
        self.options = webdriver.ChromeOptions()
        self.options.add_experimental_option("prefs", {
        #ダウンロード先のフォルダ
        "download.default_directory" : downloadDirectory, 
        #PDFをブラウザのビューワーで開かせない
        "plugins.always_open_pdf_externally": True
        })

        # ドライバーパス
        self.driverPath = driverPath

        # アクセスURL
        self.accessUrl = accessUrl

        # ログインID
        self.loginId = loginId

        # ログインPASS
        self.loginPass = loginPass

        # logger設定
        self.log = logging.getLogger('main').getChild("RecruitWebSite")

    # 採用者情報ダウンロード
    def recruitInfoDownload(self) -> None:
        try:
            # Chrome起動
            driver = webdriver.Chrome(self.driverPath, options=self.options)
            self.log.debug("Chrome起動 - 完了")

            # Googleにアクセス
            driver.get("https://aura-mico.jp/login")
            self.log.debug("Googleにアクセス - 完了")
            time.sleep(3)

            # カレントウインドウを最大化する
            driver.maximize_window()
            self.log.debug("カレントウインドウを最大化する - 完了")
            time.sleep(3)

            # ユーザIDを入力
            driver.find_element(By.ID, "exampleInputEmail1").send_keys("reworks@aura")
            self.log.debug("ユーザIDを入力 - 完了")
            time.sleep(3)

            # パスワードを入力
            driver.find_element(By.ID, "exampleInputPassword1").send_keys("hrprime")
            self.log.debug("パスワードを入力 - 完了")
            time.sleep(3)

            # ログインボタンをクリック
            driver.find_element(By.XPATH, "//*[@id='loginform4recaptcha']/button").click()
            self.log.debug("ログインボタンをクリック - 完了")
            time.sleep(7)

            # お知らせが表示されなかった場合エラーでキャッチ
            try:
                # 親フレームからiframeに移動
                iframe = driver.find_element(By.XPATH, "//*[@id='intercom-modal-container']/iframe")
                driver.switch_to.frame(iframe)
                time.sleep(3)

                # お知らせポップアップの「×」ボタンをクリック
                driver.find_element(By.XPATH, "//*[@id='intercom-container']/div/span/div/div/div/div[2]/div/div/span").click()
                self.log.debug("お知らせポップアップの「×」ボタンをクリック - 完了")
                time.sleep(3)

                # iframeから親フレームに移動
                driver.switch_to.default_content()
                time.sleep(3)
            except NoSuchElementException as e:
                self.log.debug("お知らせが表示されなかった模様")
                self.log.exception(e)

            # カテゴリーをクリック
            driver.find_element(By.XPATH, "/html/body/div[1]/header/div[2]/div/div[1]").click()
            self.log.debug("カテゴリーをクリック - 完了")
            time.sleep(3)

            # 採用をクリック
            driver.find_element(By.XPATH, "/html/body/div[1]/header/div[2]/div/div[1]/div[2]/div/div/ul/li[6]").click()
            self.log.debug("採用をクリック - 完了")
            time.sleep(3)

            # 実行をクリック
            driver.find_element(By.XPATH, "/html/body/div[1]/div[2]/div[1]/div/div[3]/button[2]").click()
            self.log.debug("実行をクリック - 完了")
            time.sleep(3)

            # ユーザー一覧をクリック
            driver.find_element(By.XPATH, "/html/body/div[1]/aside/ul/li[2]/ul/li[1]").click()
            self.log.debug("ユーザー一覧をクリック - 完了")
            time.sleep(3)

            # 全件チェックをクリック
            driver.find_element(By.XPATH, "//*[@id='check-all-student']").click()
            self.log.debug("全件チェックをクリック - 完了")
            time.sleep(3)

            # 実行をクリック
            driver.find_element(By.XPATH, "//*[@id='checked-confirm']/div/div[3]/button[2]").click()
            self.log.debug("実行をクリック - 完了")
            time.sleep(3)

            # アクションをクリック
            driver.find_element(By.XPATH, "/html/body/div[1]/div[4]/div[1]/div[1]/div[1]/div[1]/div/div[1]").click()
            self.log.debug("アクションをクリック - 完了")
            time.sleep(3)

            # CSVダウンロードをクリック
            driver.find_element(By.XPATH, "/html/body/div[1]/div[4]/div[1]/div[1]/div[1]/div[1]/div/div[2]/div/div/ul/li[5]").click()
            self.log.debug("CSVダウンロードをクリック - 完了")
            time.sleep(3)

            # なしをクリック
            driver.find_element(By.XPATH, "//*[@id='modal-export-student']/div[1]/div/div[2]/div/div[2]/div[2]/span").click()
            self.log.debug("なしをクリック - 完了")
            time.sleep(3)

            # 採用_新をクリック
            driver.find_element(By.XPATH, "//*[@id='modal-export-student']/div[1]/div/div[2]/div/div[2]/div[2]/span/div/div/ul/li[2]").click()
            self.log.debug("採用_新をクリック - 完了")
            time.sleep(3)

            # ページの最下部に移動
            driver.find_element(By.XPATH, "/html/body").send_keys(Keys.PAGE_DOWN)
            self.log.debug("ページの最下部に移動 - 完了")
            time.sleep(3)

            # ダウンロードをクリック
            driver.find_element(By.XPATH, "//*[@id='modal-export-student']/div[1]/div/div[2]/div/div[7]/button").click()
            self.log.debug("ダウンロードをクリック - 完了")
            time.sleep(3)

            # ダウンロードをクリック
            driver.find_element(By.XPATH, "//*[@id='vue-component']/div/div[4]/div/div[2]/div[1]/div/div[3]/button[2]").click()
            self.log.debug("ダウンロードをクリック - 完了")
            time.sleep(60)

            # データ作成を完了をクリック
            driver.find_element(By.XPATH, "//*[@id='vue-component']/div/div[1]/div/div/div[3]/div[1]/div[2]/table/tbody/tr[1]/td[5]/div/button").click()
            self.log.debug("データ作成を完了をクリック - 完了")
            time.sleep(3)

        except Exception  as e:
            self.log.exception(e)
            raise

        finally:
            # Chromeを終了
            driver.quit()
            self.log.debug("Chromeを終了 - 完了")
