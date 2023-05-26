import pygsheets
import google.auth
from google.oauth2 import service_account
from datetime import datetime
import pysnooper


class ReadGSheet:
    def __init__(self):
        # setting https://console.developers.google.com/
        self.credentials = 'credentials.json'  # Json 的單引號內容請改成妳剛剛下載的那個金鑰
        # 配置 Google Sheets API 凭据
        credentials = service_account.Credentials.from_service_account_file(
            self.credentials,
            scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        )
        # 使用凭据创建客户端
        self.gc = pygsheets.authorize(custom_credentials=credentials)
        self.mydata = self.read_txt_to_dict(r"data.txt")

    @staticmethod
    def read_txt_to_dict(file_name):
        # 開啟文件並讀取內容
        with open(file_name, "r", encoding="utf-8") as file:
            lines = file.readlines()

        result = {}
        for line in lines:
            # 分割每行的資訊並加入字典
            key, value = line.strip().split(":")
            result[key] = value

        return result

    def gsheet_cover(self, spreadsheet_id):
        sheet = self.gc.open_by_key(spreadsheet_id).sheet1
        # 从 Google Sheets 读取数据
        data = sheet.get_all_values()

        # 将数据转换为字典
        dict_x = {data[i][0]: data[i + 1][0] for i in range(0, len(data), 2)}

        # 打印字典
        print(dict_x)
        return dict_x

    @staticmethod
    def line_notify(msg, token):
        import requests

        LINE_NOTIFY_TOKEN = token
        print(LINE_NOTIFY_TOKEN)

        url = 'https://notify-api.line.me/api/notify'
        headers = {
            'Authorization': 'Bearer ' + LINE_NOTIFY_TOKEN  # 設定權杖
        }
        data = {
            'message': msg  # 設定要發送的訊息
        }

        return requests.post(url, headers=headers, data=data)  # 發送 LINE Notify

    # @pysnooper.snoop()
    def check_booking(self):
        def record_sent_date(date):
            with open('sent_dates.txt', 'w') as f:
                f.write(str(date))

        def get_sent_dates():
            try:
                with open('sent_dates.txt', 'r') as f:
                    return f.read()  # 返回整个文件内容作为一个字符串
            except FileNotFoundError:
                return ''

        gsheet_cover = self.mydata["gsheet_cover"]
        sheet = self.gc.open_by_key(gsheet_cover).sheet1
        # 从 Google Sheets 读取数据
        cell_value = sheet.get_value("A2")

        # 将数据转换为列表
        date_l = [int(x) for x in cell_value.split('/')]

        # 打印列表
        print(date_l)

        # 当前日期
        today = datetime.today()

        # 将列表转换为 datetime 对象
        given_date = datetime(today.year, date_l[0], date_l[1])

        # 比较今天的日期与给定的日期
        if given_date >= today:
            job = "Bus"
            print("给定的日期是未来时间")
            from update_jenkins_job import UpdateJenkinsJob
            jenkins = UpdateJenkinsJob()
            # spec.text = "H(58-59) 6 30 4 *"

            job_up_time = f"58 6 {date_l[1]} {date_l[0]} *"
            print(f"update [{job}] => {job_up_time}")
            jenkins.job_update_trigger(job, job_up_time)

            # <<<<<<< HEAD
            #             get_date = self.get_sent_dates()
            #             date_string = f"{date_l[0]}/{date_l[1]}"
            #             if get_date != date_string:
            #                 self.line_notify(f"Job={job}, Set trigger date: {date_l[0]}/{date_l[1]}, AM:6:58",
            #                                  token=self.mydata["line_token"])
            #                 self.record_sent_date(date_string)
            # =======
            sent_dates = get_sent_dates()
            formatted_date = f"{date_l[0]}/{date_l[1]}"

            if sent_dates != formatted_date:
                self.line_notify(f"Job={job}, Set trigger date: {date_l[0]}/{date_l[1]}, AM:6:58",
                                 token=self.mydata["line_token"])

            self.record_sent_date(formatted_date)

        # >>>>>>> feature
        else:
            print("给定的日期是过去时间")

    @staticmethod
    def record_sent_date(date):
        with open('sent_dates.txt', 'w') as f:
            f.write(str(date))

    @staticmethod
    def get_sent_dates():
        try:
            with open('sent_dates.txt', 'r') as f:
                return f.read()
        except FileNotFoundError:
            return []


if __name__ == '__main__':
    read_gc = ReadGSheet()
    read_gc.check_booking()
