import jenkins


class UpdateJenkinsJob:
    def __init__(self):
        # 連接到 Jenkins 服務器
        jenkins_url = "http://10.0.0.3:18081"

        mydata = self.read_txt_to_dict(r"data.txt")
        jenkins_username = mydata["jenkins_name"]
        jenkins_password = mydata["jenkins_passwd"]
        self.server = jenkins.Jenkins(jenkins_url, username=jenkins_username, password=jenkins_password)

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

    def job_update_trigger(self, job_name, job_time):
        # 讀取指定 job 的配置
        job_config = self.server.get_job_config(job_name)

        # 使用 lxml 解析 XML 並更改 Build Triggers
        from lxml import etree

        # root = etree.fromstring(job_config)
        root = etree.fromstring(job_config.encode('utf-8'))

        # 查找 triggers 部分
        triggers = root.find("triggers")

        # 假設您想要添加一個定時觸發器
        timer_trigger = etree.Element("hudson.triggers.TimerTrigger")
        spec = etree.Element("spec")
        # spec.text = "H(58-59) 6 30 4 *"
        spec.text = job_time
        timer_trigger.append(spec)

        # 將新的觸發器添加到 triggers 節點
        triggers.append(timer_trigger)

        # 更新 Jenkins job 配置
        new_job_config = etree.tostring(root, pretty_print=True).decode("utf-8")
        self.server.reconfig_job(job_name, new_job_config)


if __name__ == '__main__':
    jenkins = UpdateJenkinsJob()
    jenkins.job_update_trigger("Buy_Amai", "H(58-59) 6 22 2 *")
