import pygsheets
import google.auth
from google.oauth2 import service_account
from datetime import datetime
import pysnooper
from gmail_sender import GmailSender


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
        # 初始化 Gmail 寄件器
        self.gmail_sender = GmailSender(
            sender_email=self.mydata["gmail_sender"],
            app_password=self.mydata["gmail_password"]
        )

    @staticmethod
    def read_txt_to_dict(file_name):
        # 開啟文件並讀取內容
        with open(file_name, "r", encoding="utf-8") as file:
            lines = file.readlines()

        result = {}
        for line in lines:
            # 分割每行的資訊並加入字典
            key, value = line.strip().split(":")
            # 如果是收件者列表，則分割成列表
            if key == "recipient_emails":
                result[key] = value.split(",")
            else:
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

    def send_email(self, msg, email_config):
        try:
            return self.gmail_sender.send_email(
                recipient_emails=self.mydata["recipient_emails"],
                subject="Jenkins Job 更新通知",
                text_content=msg,
                sender_name=self.mydata.get("name", "預約系統通知")
            )
        except Exception as e:
            print(f"郵件發送失敗：{str(e)}")
            return False

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
            job = "ycbus"
            print("给定的日期是未来时间")
            from update_jenkins_job import UpdateJenkinsJob
            jenkins = UpdateJenkinsJob()

            job_up_time = f"00 7 {date_l[1]} {date_l[0]} *"
            print(f"update [{job}] => {job_up_time}")
            jenkins.job_update_trigger(job, job_up_time)

            sent_dates = get_sent_dates()
            formatted_date = f"{date_l[0]}/{date_l[1]}"

            if sent_dates != formatted_date:
                self.send_email(
                    f"Job={job}, Set trigger date: {date_l[0]}/{date_l[1]}, AM:7:00",
                    {}  # 空字典，因為我們現在直接從 self.mydata 讀取
                )
                record_sent_date(formatted_date)
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
