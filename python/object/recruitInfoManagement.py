import logging
import glob
import pandas as pd
import requests
import os

# 採用者情報管理クラスを定義
class RecruitInfoManagement:

    # コンストラクタを定義
    def __init__(self, workBook, csvDir, getSheetName, printSheetName, convertSheetName) -> None:
        # ワークブック
        self.workBook = workBook

        # CSVファイルパス
        self.csvDir = csvDir

        # 取得用ワークシート名
        self.getSheetName = getSheetName

        # 貼付用ワークシート名
        self.printSheetName = printSheetName

        # 変換用ワークシート名
        self.convertSheetName = convertSheetName

        # 採用者の番号
        self.recruitNumber = 0

        # 登録対象件数
        self.recruitCount = 0

        # logger設定
        self.log = logging.getLogger('main').getChild("RecruitInfoManagement")

    # 採用者登録
    def RegisterRecruitInfo(self) -> None:
        try:
            self.getJobInformation()
            self.getRecruitNumber()
            self.duplicateDeletion()

            # 登録対象者が０件の場合は処理しない
            if(self.recruitCount > 0):
                self.printRecruitInfo()
                self.postData()
        except Exception as e:
            raise e

        try:
            # CSVファイル削除
            self.log.debug("CSVファイル削除 - 開始")
            for file in glob.glob(self.csvDir):
                os.remove(file)
            self.log.debug("CSVファイル削除 - 完了")
        except Exception as e:
            self.log.error("CSVファイル削除 - 失敗")
            self.log.exception(e)


     # 採用情報取得
    def getJobInformation(self) -> None:
        try:
            self.log.debug("採用情報取得 - 開始")
            self.jobInformation = pd.read_csv(glob.glob(self.csvDir)[0], encoding="cp932").fillna("")
            self.jobInformation = self.jobInformation.values.tolist()
            self.log.debug("読み込んだファイル：" + glob.glob(self.csvDir)[0])
            self.log.debug("採用情報取得 - 完了")
        except IndexError as e:
            self.log.error("採用情報取得 - 失敗")
            self.log.error("CSVファイルを発見できなかった模様")
            self.log.exception(e)
            raise e
        except Exception as e:
            self.log.error("採用情報取得 - 失敗")
            self.log.exception(e)
            raise e

    # 前回の実行時に登録した採用者の番号を取得する
    def getRecruitNumber(self) -> None:
        try:
            self.log.debug("採用者番号取得 - 開始")
            self.recruitNumber = self.workBook.worksheet(self.getSheetName).cell(1, 1).value
            self.log.debug("前回の実行時に登録した採用者の番号：" + self.recruitNumber)
            self.log.debug("採用者番号取得 - 完了")
        except Exception as e:
            self.log.error("採用者番号取得 - 失敗")
            self.log.exception(e)
            raise e

    # 重複削除
    def duplicateDeletion(self) -> None:
        try:
            self.log.debug("重複削除 - 開始")
            index = 1
            for list in self.jobInformation:

                # 前回の実行時に登録した採用者の番号と一致するか判定
                if(int(list[0]) == int(self.recruitNumber)):
                    break

                index = index + 1

            # 前回の実行時に登録していた採用者情報を削除(一人もいない場合は実行しない)
            if(index > 1):
                del self.jobInformation[0:index]

            # 登録対象の件数を取得
            self.recruitCount = len(self.jobInformation)
            self.log.debug("登録対象件：" + str(self.recruitCount) + "件")
            self.log.debug("重複削除 - 完了")
        except Exception as e:
            self.log.error("重複削除 - 失敗")
            self.log.exception(e)
            raise e

    # データ貼り付け
    def printRecruitInfo(self) -> None:
        try:
            self.log.debug("データ貼り付け - 開始")
            # 貼り付ける前にシートをクリア
            self.workBook.worksheet(self.printSheetName).clear()

            self.workBook.values_append(self.printSheetName, {"valueInputOption" : "USER_ENTERED"}, {"values" : self.jobInformation})
            self.log.debug("データ貼り付け - 完了")
        except Exception as e:
            self.log.error("データ貼り付け - 失敗")
            self.log.exception(e)
            raise e

    # GAS起動
    def postData(self):
        try:
            self.log.debug("GAS起動 - 開始")
            # パラメータ設定 
            url = "https://script.google.com/macros/s/AKfycbzTd0Y_gyxe2HGt-DBQkylacGUS8LNviw7mTaO4AAgx552xvfEigIQCHJW9uZRd5Iv4/exec"
            headers = {'Content-Type': 'application/json',}
            response = requests.post(url, data="", headers=headers)

            # 実行結果ステータス判定
            if(response.status_code == 200 and response.text == "success"):
                self.log.debug(response.text)

                # 成功した場合のみ採用者の番号を更新
                self.recruitNumber = self.workBook.worksheet(self.getSheetName).update_cell(1, 1, self.jobInformation[self.recruitCount - 1][0])

            else:
                self.log.error(response.text)
                raise e
            self.log.debug("GAS起動 - 完了")
        except Exception as e:
            self.log.error("GAS起動 - 失敗")
            self.log.exception(e)
            raise e