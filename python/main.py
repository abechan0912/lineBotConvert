import logging
from logging import config
import yaml
from oauth2client.service_account import ServiceAccountCredentials
import gspread
from object import recruitWebSite
from object import recruitInfoManagement
import sys

# logger設定
config.fileConfig("C:\\programing\\python\\rpa\\lineBotConvert\\resource\\config\\logging.conf")
log = logging.getLogger('main')

# パラメータファイル
with open("C:\\programing\\python\\rpa\\lineBotConvert\\resource\\config\\parameters.yaml", encoding="utf-8") as file:
    para = yaml.safe_load(file)

    # MicoCloud
    micoCloud = recruitWebSite.RecruitWebSite(
        para["CHROME"]["DOWNLOAD_DIRECTORY"], 
        para["CHROME"]["DRIVER_PATH"], 
        para["MICO_CLOUD"]["ACCESS_URL"], 
        para["MICO_CLOUD"]["LOGIN_ID"], 
        para["MICO_CLOUD"]["LOGIN_PASS"]
    )

    # 採用者情報ダウンロード
    try:
        log.debug("採用者情報ダウンロード - 開始")
        micoCloud.recruitInfoDownload()
        log.debug("採用者情報ダウンロード - 完了")
    except Exception as e:
        log.error("採用者情報ダウンロード - 失敗")
        sys.exit()

    # Google認証
    try:
        log.debug("Google認証 - 開始")
        # API設定
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

        # 認証情報設定
        credentials = ServiceAccountCredentials.from_json_keyfile_name(para["COMMON"]["CERTIFICATION_KEY"], scope)
        gc:any = gspread.authorize(credentials)

        # スプレッドシート取得
        workBook = gc.open_by_key(para["COMMON"]["SPREADSHEET_KEY"])
        log.debug("Google認証 - 完了")
    except Exception as e:
        log.error("Google認証 - 失敗")
        log.exception(e)
        sys.exit()

    # MicoCloud 採用者情報
    micoCloudInfo = recruitInfoManagement.RecruitInfoManagement(
        workBook,
        para["GOOGLE_SHEETS"]["CSV_DIR"] + para["GOOGLE_SHEETS"]["CSV_NAME"], 
        para["GOOGLE_SHEETS"]["GET_SHEET_NAME"], 
        para["GOOGLE_SHEETS"]["PRINT_SHEET_NAME"],
        para["GOOGLE_SHEETS"]["CONVERT_SHEET_NAME"]
    )

    # 採用者登録
    try:
        log.debug("採用者登録 - 開始")
        micoCloudInfo.RegisterRecruitInfo()
        log.debug("採用者登録 - 完了")
    except Exception as e:
        log.error("採用者登録 - 失敗")
        sys.exit()