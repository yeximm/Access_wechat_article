import requests
import os
import re       # 使用正则表达式
import json     # 用于json转码
from bs4 import BeautifulSoup


class SaveContent():
    def __init__(self):
        self.save_root_path = r'./data'
        self.names = '公众号默认名称'      # 程序默认以‘公众号名称’做目录，用以保存相关公众号的信息
        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) '
                          'Chrome/116.0.0.0 Safari/537.36 NetType/WIFI MicroMessenger/7.0.20.1781(0x6700143B) '
                          'WindowsWechat(0x63090a13) XWEB/9129 Flue',
        }
        self.cookies = {
            "poc_sid": '',      # 用以保存设备ID
        }
        os.makedirs(self.save_root_path, exist_ok=True)  # 创建保存路径，如果文件夹已存在，则忽略，默认为r'./data'

    def get_all_content(self, content):
        """
            汇总所有的信息
            :param content:网页内容
        """
        title = self.get_title(content)
        times = self.get_time(content)
        texts = self.get_texts(content)
        return {
            'title': title,     # 文章标题
            'times': times,     # 文章发布时间
            'texts': texts,     # 文章内容（列表形式）
            'all_type': '',  # 文章包含的标志信息，用于获取文章阅读量等信息
        }

    def verify_user(self, url, content):
        """
            url=请求路径
            content=网页内容，如：res.text
            遇到此情况时使用：  >当前环境异常，完成验证后即可继续访问。<
            返回参数；验证标志（1为有效），网页内容，cookie值
                {'verify_flag': 1, 'content': res.text, 'poc_sid': poc_sid}
        """
        print('开始验证，正在获取参数poc_sid')
        poc_token = re.search(r'poc_token.*"(.*?)"', content).group(1)
        poc_sid = re.search(r'poc_sid.*"(.*?)"', content).group(1)  # poc_sid为cookie参数
        cap_appid = re.search(r'cap_appid.*"(.*?)"', content).group(1)
        cap_sid = re.search(r'cap_sid.*"(.*?)"', content).group(1)
        target_url = re.search(r'target_url.*"(.*?)"', content).group(1)

        try:
            '''验证请求第一步'''
            verify1_url = ('https://t.captcha.qq.com/cap_union_prehandle?'
                           'aid=' + cap_appid +
                           '&protocol=https&accver=1&showtype=popup&ua=TW96aWxsYS81LjAgKFdpbmRvd3MgTlQgMTAuMDsgV2luNjQ7IHg2NCkgQXBwbGVXZWJLaXQvNTM3LjM2IChLSFRNTCwgbGlrZSBHZWNrbykgQ2hyb21lLzEyNy4wLjAuMCBTYWZhcmkvNTM3LjM2IEVkZy8xMjcuMC4wLjA%3D&noheader=0&fb=1&aged=0&enableAged=0&enableDarkMode=1&'
                           'deviceID=' + poc_sid + '&sid=' + cap_sid +
                           '&grayscale=1&dyeid=0&clientype=2&cap_cd=&uid=&lang=zh-cn&entry_url=https%3A%2F%2Fmp.weixin.qq.com%2Fmp%2Fwappoc_appmsgcaptcha&elder_captcha=0&js=%2Ftcaptcha-frame.8d77d8b0.js&login_appid=&wb=1&version=1.1.0&subsid=1&callback=_aq_873604&sess=')
            verify1 = requests.get(verify1_url, headers=self.headers, verify=False)
            # sess，pow_answer 为第三步中的请求数据
            sess = re.search('sess":"(.*?)"', verify1.text).group(1)
            pow_answer = re.search('prefix":"(.*?)"', verify1.text).group(1)

            '''验证请求第二步'''
            verify2_url = 'https://t.captcha.qq.com' + re.search('tdc_path":"(.*?)"', verify1.text).group(1)
            verify2 = requests.get(verify2_url, headers=self.headers, verify=False)
            # eks 为第三步中的请求数据
            eks = re.search(r"='(.*?)'", verify2.text).group(1)

            '''验证请求第三步'''
            verify3_url = 'https://t.captcha.qq.com/cap_union_new_verify'
            verify3_data = {
                # 'collect': 'rTdL2csh9MFPSaeKi0kIljgBgCZKq9QgnKBxrlpTn3YT41x6hx3sIjZVFn544i8YgxUt4iyQSxZpkD%2BeFzrQB4MzlnMqochzx0%2BrFwXMMEnmz5ec2RysL1qlKV9BAQz5K3gw7jWpBPJWtyWKD5UWTK39jnbMtitCUGRFdQQS0XYeIX3GYdrmSJc0a4JdPx3YuZW5YTdar%2BhbG07vWc8IsCWv9WYfmFkfvXzUWm5alHEuoKngABAFUVZWlvaulCpZ89U5cI90c9Cuz4gVc5q3bXOTNk1hWo3dF3CfnbVnoezWefsjIzyR9Fn1fNK2nv7rVMYUdPORoo5WVpb2rpQqWVZWlvaulCpZC5095cBnxOUbkZU9ko3C0CF3qrZy0qXgg%2BcsfLgi6mn63I2eOkQD%2FISjcH3WeTtN37dGnsQOeLGofkN2MCNpMirzT2sFN4ZOImNS%2FY6FQOGZ7UQsyRg8xly%2FO5ChYYsBMIru78AYuJi2tACDtOxCi0Sepo7US8w8LwtCIyNOUQQO4531ay7Hrexsw4KgeneWb2cj%2FLzXypA2Vga97ex9S61bW%2B1BAHTpDglMOA5VC5VbrVAau6mh7HkWP90YufkLAfKedGQRWpTyXiDJlKhbu5OC3ZAcJXDtHvI0uQtGD6j4y0edGMc0bT9zDNW33WLsL%2BEVpVv7AuHl9VlmZ0Zc5Q8BO0yDe70FQgIpI5F%2BWoKhUxQDqAtbL2WeCGqbj9AG9PAZkMr620aAsIGBSCufUG22gnWrAFUJ6okHj5Q7kYycBYBoYjkFkv8WAof1Z5WkOkwQ%2FUGx7NyOY5VyKZPqEcj5n2r7r4rL2Ha35M23XH73CDno6e6QHeSXsYdUw2%2BG4l7Rp8QNyAlzzVeAP69v7oZ%2BjuAxxmLBwEtFRk0qVnlW%2F%2BZYj7LybHtpL%2FSt9V2egQiO1Y5qZ1z36gbbnborLA2%2BEQIQYd01SZnsl7OsO0cViJ4wBEcfgCOMBDvnJBfyenM1nRY3zaeuecnhbJIzK33ROtYsD3Y7o%2FKYghPdictG32I59qyo5qn2gG4aPaWoTOOBa9jQU3nJy6TRgtpw1p1zsthUoo%2B3Z9YxZlYEJZXdsG2797jLExc7vg3BwXxKfxIUezrobzkABVVH%2FTYkNnraXUv8GAtvAmFdyPK7y%2FBYzBUw9SfmYGAvemDFbjc0kesd2Pl7rME51FURZdu8AW%2Bni6reapdxg4u5REEfUaZbXT7bE4aqwIMPqcezvluHIghv3dCZZ0n%2Fn1zNADEjB1ZWlvaulCpZcg0Mp9p9ZhTR9MfLjjO6to6e2sUK9OCUkI6fIkvjG0ooDtXbCdfD5Q%3D%3D',
                # 'tlg': 1336,
                'eks': eks,
                'sess': sess,
                'ans': '[{"elem_id":0,"type":"DynAnswerType_TIME","data":""}]',
                'deviceID': poc_sid,
                'pow_answer': pow_answer,
                # 'pow_calc_time': 3,
            }
            verify3 = requests.post(verify3_url, headers=self.headers, data=verify3_data, verify=False)
            # ticket，randstr为第四步请求参数
            ticket = json.loads(verify3.text)['ticket']
            randstr = json.loads(verify3.text)["randstr"]

            '''验证请求第四步'''
            verify4_url = 'https://mp.weixin.qq.com/mp/wappoc_appmsgcaptcha?action=Check&x5=0&f=json'
            verify4_data = {
                'target_url': target_url,
                'poc_token': poc_token,
                'appid': cap_appid,
                'ticket': ticket,
                'randstr': randstr,
            }
            self.cookies['poc_sid'] = poc_sid   # 重置类属性 cooikes的值
            print('获取到poc_sid：'+ self.cookies['poc_sid'])
            verify4 = requests.post(verify4_url, headers=self.headers, cookies=self.cookies, data=verify4_data, verify=False)
            # print(verify4.text)
            # print('发送成功后，poc_sid就可以正常使用了')

            '''验证请求第五步'''
            modify_url = url + '&poc_token=' + poc_token
            res = requests.get(modify_url, headers=self.headers, cookies=self.cookies, verify=False)  # 发起请求
            print('已完成验证请求，后续请求可正常使用。')
            return {'verify_flag': 1, 'content': res.text, 'poc_sid': poc_sid}
        except:
            print('验证失败，请检查后再进行尝试')
            return {'verify_flag': 0}

    def get_content(self, url):
        """
            输入微信文章链接（永久链接或临时链接）
            返回参数：请求标志位（1时有效），文章内容（网页内容），请求参数cookie
            {'content_flag': 1, 'content': res.text, 'poc_sid': self.cookies['poc_sid']}
        """
        res = requests.get(url, headers=self.headers, cookies=self.cookies, verify=False)  # 发起请求
        # 验证请求
        if 'var createTime = ' in res.text:     # 正常获取到文章内容
            print('正常获取到文章内容')
            self.names = re.search(r'var nickname.*"(.*?)".*', res.text).group(1)  # 公众号名称
            return {'content_flag': 1, 'content': res.text, 'poc_sid': self.cookies['poc_sid']}
        elif '>当前环境异常，完成验证后即可继续访问。<' in res.text:
            print('当前环境异常，完成验证后即可继续访问。')
            # 返回环境异常，需进行验证操作
            contents = self.verify_user(url, res.text)
            if contents['verify_flag'] == 1:
                self.names = re.search(r'var nickname.*"(.*?)".*', contents['content']).group(1)  # 公众号名称
                return {'content_flag': 1, 'content': contents['content'], 'poc_sid': contents['poc_sid']}
            else:
                print('验证操作失败！！！！！请检查后重试')
                return {'content_flag': 0}
        elif '操作频繁，请稍后再试。' in res.text:
            # 锁cookie了，等待几个小时后再尝试
            print('操作频繁了，等会再弄或换ip弄！！！')
            print(res.text)
            return {'content_flag': 0}
        else:   # 出现错误信息，待处理！！！！！！！！！
            print("请求出错，请等待几秒后请继续操作")
            print('出现其他问题，请查找原因后再试！！！！')
            return {'content_flag': 0}

    def get_img(self, content, title, detail_time):
        """
            将文章中的图片下载，并存储到本地
            存储路径在 self.save_root_path 下
            搭配create_directory_and_download_image()使用
        """
        imgs = content.split('https://mmbiz.qpic.cn/')
        # 在Windows操作系统中，文件名需要遵守以下命名规则：
        # 文件名不包含以下任何字符：“（双引号）、*（星号）、<（小于）、>（大于）、？（问号）、\（反斜杠）、/（正斜杠）、|（竖线）、：（冒号）。
        # 文件名不要以空格、句号、连字符或下划线开头或结尾。
        title = title.replace('"', '“').replace('*', '※').replace('|', 'I').replace('/', ' ')
        title = title.replace('\\', ' ').replace(':', '：').replace('<', '《').replace('>', '》').replace('?', '？')
        img_save_path = self.names + "/" + detail_time[0] + '_' + title    # 图片保存路径

        for i in range(0, len(imgs) - 1):
            img_url = 'https://mmbiz.qpic.cn/' + imgs[i + 1].split('"')[0]
            # print('正在获取图片：' + img_url)
            self.create_directory_and_download_image(img_url, self.save_root_path + "/" + img_save_path, str(i + 1))

    def create_directory_and_download_image(self, image_url, directory_path,img_name):
        image_name = ''
        os.makedirs(directory_path, exist_ok=True)
        response = requests.get(image_url, cookies=self.cookies, verify=False)
        if response.status_code == 200:
            img_hz = ['gif', 'jpg']
            for imghz in img_hz:
                if imghz in image_url:
                    image_name = img_name + '.' + imghz
            if image_name == '':
                image_name = img_name + '.jpg'
            file_path = directory_path + '/' + image_name
            with open(file_path, 'wb') as f:
                f.write(response.content)
            print(f"已成功下载图片： {file_path}")
        else:
            print(f"无法下载图片，状态码: {response.status_code}")

    def get_title(self, content):
        """
            文章标题在<h1></h1>标签内
            '.' 匹配任意字符，除了换行符，当re.DOTALL标记被指定时，则可以匹配包括换行符的任意字符。
            re.search(path, content, re.DOTALL).group(1)获取的内容为
                class="rich_media_title " id="activity-name">title
            split('>')将获取到的内容从'>'开始切片
            strip()删除字符串两边空白
        """
        path = r'<h1(.*)</h1'
        tittle = re.search(path, content, re.DOTALL).group(1).split('>')[1].strip()
        return tittle

    def get_time(self, content):
        """
            尝试获取时间，程序到此处时已经过验证，可以得到时间信息
            timex为正则表达式筛选出的数据：2024-08-01 07:09
            返回参数（例）： ['2024-08-01', '07:09']
        """
        timex = re.search(r'var createTime = (.*?) .*', content).group().split("'")[1]
        year = str(timex.split(' ')[0].split('-')[0])
        month = str(timex.split(' ')[0].split('-')[1])
        day = str(timex.split(' ')[0].split('-')[2])
        hour = str(timex.split(' ')[1])
        # detail_time = str(year) + '年' + str(month) + '月' + str(day) + '日'
        detail_time = [timex.split(' ')[0], hour]
        return detail_time

    def get_texts(self, content):
        """
            将文字内容转换为列表形式存储
        """
        soup = BeautifulSoup(content, 'html.parser')
        spans = soup.find_all('p')
        texts = [span.text for span in spans]
        # print('正在读取页面内容：' + str(texts))
        return texts



