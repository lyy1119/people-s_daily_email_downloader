import requests
from datetime import datetime, timedelta
import os
from PyPDF2 import PdfMerger
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import platform
# import glob

# date = "2024-08/09/"
# num = 1
# http://paper.people.com.cn/rmrb/images/ 2024-08/09/ 08 /rmrb2024080908.pdf

max_files = 20      # 最多页数
max_attempts = 5    # 最多尝试次数
today_files = 0    # 今天最多的页数

# smtp
smtp_server = 'smtp.163.com'
smtp_port = 465     # 使用ssl
# smtp_port = 587    # 使用tls
smtp_user = ''
smtp_password = ''
# 发件人和收件人
from_email = ''
recive_email = ['']

# 逻辑功能函数
def get_paper_date():
    now = datetime.now()
    current_hour = now.hour
    if current_hour < 7:
        adjusted_date = now - timedelta(days=1)     # 如果是七点前则用前一天的日期，用于测试。实际使用时应当定时执行
    else:
        adjusted_date = now
    return adjusted_date

def get_date():     # 获取日期
    adjusted_date = get_paper_date()
    formatted_date = adjusted_date.strftime('%Y-%m/%d/')
    date_number = adjusted_date.strftime('%Y%m%d')
    return formatted_date , date_number


def gen_email_body():
    now = datetime.now()

    # 获取当前小时数
    current_hour = now.hour

    # 根据小时数输出相应的问候语
    if 5 <= current_hour < 12:
        greeting = "早上好"
    elif 12 <= current_hour < 13:
        greeting = "中午好"
    elif 14 <= current_hour < 18:
        greeting = "下午好"
    else:
        greeting = "晚上好"

    body = greeting + ",现在是" + str(now.hour) + "点，今天的人民日报有" + str(today_files) + "版，详细内容请查看附件。\n" + "发送时间" + f"{now.year}年{now.month}月{now.day}日{now.hour}时{now.minute}分"

    return body

def gen_email_subject():
    now = get_paper_date()
    subject = f"{now.year}年{now.month}月{now.day}日 人民日报 邮件速递"
    return subject


# 业务功能函数
def download_all_page(logfile):
    date , date_number = get_date()       # 获取日期
    for i in range(1 , max_files + 1): # 循环下载
        # 构造完整的URL
        page_number = str(i).zfill(2)
        url = "http://paper.people.com.cn/rmrb/images/" + date + page_number + "/rmrb" + date_number + page_number + ".pdf"
        # print(url)

        # 尝试下载
        success = False
        for attempt in range(max_attempts):
            try:
                response = requests.get(url, timeout=50)
                response.raise_for_status()  # 如果请求失败，抛出HTTPError
                success = True
                break  # 下载成功则跳出重试循环
            except (requests.HTTPError, requests.RequestException) as e:
                logfile.write(f"Attempt {attempt + 1} failed for {url}: {e}\n")
        if success:
            # 保存下载的PDF
            global today_files
            today_files = i     # 记录页数
            with open(f"{page_number}.pdf", 'wb') as f:
                f.write(response.content)
            logfile.write(f"Downloaded: {url}\n")
        else:
            logfile.write(f"File {url} does not exist or could not be downloaded after {max_attempts} attempts.\n")


def merge_pdf(logfile):
    merger = PdfMerger()
    # 合并存在的PDF文件
    for i in range(1, max_files + 1):
        filename = f"{i:02d}.pdf"
        if os.path.exists(filename):
            merger.append(filename)
        else:
            logfile.write(f"File {filename} does not exist, skipping.\n")

    # 保存合并后的PDF
    now = get_paper_date()
    output_filename = f"{now.year}.{now.month}.{now.day}_people's_daily.pdf"
    merger.write(output_filename)
    merger.close()
    logfile.write(f"PDF files merged into {output_filename}\n")
    return output_filename

def send_email(logfile , pdf_name , to_email):
    subject = gen_email_subject()
    body = gen_email_body()

        # 创建MIMEMultipart对象
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    # 添加附件
    filename = pdf_name
    if os.path.exists(filename):
        with open(filename, "rb") as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header(
                'Content-Disposition',
                f'attachment; filename= {filename}',
            )
            msg.attach(part)
    else:
        logfile.write(f"File {filename} does not exist.\n")

    # 连接到SMTP服务器并发送邮件
    # 初始化server变量
    server = None
    # 连接到SMTP服务器并发送邮件
    try:
        # ssl发送
        # server = smtplib.SMTP_SSL(smtp_server, smtp_port)
        # server.login(smtp_user, smtp_password)
        # server.sendmail(from_email, to_email, msg.as_string())
        # tls发送
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()  # 开启TLS加密
        server.login(smtp_user, smtp_password)
        server.sendmail(from_email, to_email, msg.as_string())

        logfile.write("Email sent successfully.\n")
    except smtplib.SMTPServerDisconnected as e:
        logfile.write(f"SMTP server disconnected: {e}\n")
    except smtplib.SMTPException as e:
        logfile.write(f"Failed to send email: {e}\n")
    finally:
        # 只有在server已连接且未断开时才调用quit()
        if server and server.sock:
            try:
                server.quit()
            except smtplib.SMTPServerDisconnected:
                logfile.write("Server already disconnected, cannot quit.")


def del_temple_files(logfile):
    # 获取当前系统类型
    system_type = platform.system()
    # 判断系统类型并执行相应的删除命令
    if system_type == "Linux" or system_type == "Darwin":  # Darwin表示macOS
        # Linux或macOS
        os.system(f"rm -f *.pdf")
        logfile.write("Files deleted using 'rm' command.\n")
    elif system_type == "Windows":
        # Windows
        os.system(f"del /F *.pdf")
        logfile.write("Files deleted using 'del' command.\n")
    else:
        logfile.write("Unsupported operating system.\n")


if __name__ == "__main__":
    with open("logs.txt", "a") as logfile:
        now = datetime.now()
        logfile.write(f"===Run program===\ntime:{now.year}.{now.month}.{now.day}.{now.hour}.{now.minute}.{now.second}\n")

        logfile.write("=>Downloading single file...\n")
        download_all_page(logfile)

        logfile.write("=>Merging files...\n")
        pdf_name = merge_pdf(logfile)

        for i in recive_email:
            logfile.write(f"=>Sending email to {i} ...\n")
            send_email(logfile , pdf_name , i)

        logfile.write("=>Deleting pdf files...\n")
        del_temple_files(logfile)