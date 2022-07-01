from telegram.ext import (
    Updater,
    CommandHandler,
    CallbackQueryHandler,
    MessageHandler,
    Filters,
    ContextTypes,
    ConversationHandler
)
from telegram import (
    InlineKeyboardMarkup,
    InlineKeyboardButton,
    ReplyKeyboardMarkup,
    ReplyKeyboardRemove,
    Update,
    Bot,
)
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
from threading import Thread, current_thread
import telegram
import datetime
import imaplib
import smtplib
import email
import yaml
import time
import os
import sys


# TODO v1:
#   html 转图片发送
#       使用 API 转换 (√) 7
#       使用 imgkit 转换 (未使用)
#       使用 Telegram 自带的 Web Page 显示 (v13.13 暂不支持)
#   删除邮件 (已完成, v1.0.0)
#   支持更多的附件 3
#   yaml 配置文件解析 1
#   使用 pprint 输出日志 1
#   Docker 镜像部署 8 (v1.8.0)
#   回复邮件 3
#   修改检查邮箱延时 2
#   批量删除邮件 4
#   回复某一条邮件时提供回复、删除按钮 5
#   对回复邮件的发收件人的格式进行处理，使其可以直接在机器人窗口发信 6

# TODO v2:
#   支持多个邮箱
#   支持多个用户配置使用机器人

class Email2TGUtil:
    """
    Email2TG工具类
    """
    imap_host = ''
    imap_enable_ssl = True
    smtp_host = ''
    smtp_enable_ssl = True
    username = ''
    password = ''
    mail_box = ''
    delay_time = 5
    tg_chat_id = ''
    tg_bot_token = ''
    imap = None
    smtp = None

    def __init__(self, imap_host, imap_enable_ssl, smtp_host, smtp_enable_ssl):
        """初始化方法"""
        self.imap_host = imap_host
        self.smtp_host = smtp_host
        self.imap_enable_ssl = imap_enable_ssl
        self.smtp_enable_ssl = smtp_enable_ssl
        # 初始化邮箱对象
        try:
            if imap_enable_ssl:
                self.imap = imaplib.IMAP4_SSL(host=self.imap_host, port=993)
            else:
                self.imap = imaplib.IMAP4(host=self.imap_host, port=143)
            if smtp_enable_ssl:
                self.smtp = smtplib.SMTP_SSL(host=self.smtp_host, port=465)
            else:
                self.smtp = smtplib.SMTP_SSL(host=self.smtp_host, port=25)
            tgprint("邮件服务器已连接")
        except:
            tgprint("邮件服务器连接失败, 请检查服务器配置是否正确")
            exit(0)

    def login(self, username, password):
        """登录"""
        self.username = username
        self.password = password
        try:
            self.imap.login(user=self.username, password=self.password)
            self.smtp.login(user=self.username, password=self.password)
        except:
            tgprint("邮件服务器登录失败, 请检查账号配置是否正确")
            exit(0)

    def configure(self, delay_time, tg_chat_id, tg_bot_token):
        """信息配置"""
        self.delay_time = delay_time
        self.tg_chat_id = tg_chat_id
        self.tg_bot_token = tg_bot_token

    def send_http_mail(self, recv_email, header_text, mail_msg):
        """发送HTTP格式邮件"""
        msg = MIMEText(mail_msg, 'html', 'utf-8')
        msg['Subject'] = header_text
        msg['From'] = username
        msg['To'] = recv_email
        try:
            self.smtp.sendmail(username, recv_email, msg.as_string())
            tgprint("发送成功")
        except Exception as e:
            tgprint("发送失败, 请检查配置")
            print(e)

    def send_multipart_mail(self, recv_email, header_text, mail_msg, file_name):
        """发送Multipart(带附件)邮件"""
        msg = MIMEMultipart()
        msg['From'] = username
        msg['Subject'] = header_text
        msg['To'] = recv_email
        msg.attach(MIMEText(mail_msg, 'html', 'utf-8'))
        att1 = MIMEText(open(file_name, 'rb').read(), 'base64', 'utf-8')
        att1["Content-Type"] = 'application/octet-stream'
        att1["Content-Disposition"] = 'attachment; filename="' + file_name + '"'
        msg.attach(att1)
        try:
            self.smtp.sendmail(username, recv_email, msg.as_string())
            tgprint("发送成功")
        except Exception as e:
            tgprint("发送失败")
            tgprint(e)
        os.remove(file_name)

    def get_mail(self, mail_len):
        """获取邮件"""
        self.mail_box = mail_box
        if self.imap is not None:
            self.imap.select(readonly=False)
            typ, data = self.imap.search(None, 'ALL')
            if typ == 'OK':
                tl = data[0].split()
                if int(mail_len) > len(tl):
                    tgprint("无效的输入.\n你的邮箱内邮件数量为" + str(len(tl)) + ", 你想要的太多了.")
                    return
                for num in tl[len(tl) - int(mail_len):]:
                    typ1, data1 = self.imap.fetch(num, '(RFC822)')  # 读取邮件
                    if typ1 == 'OK':
                        self.output_mail_text(data1, num)
        else:
            tgprint("邮件服务器连接异常，正在重启 错误代码: 0x00")
            restartAuto()

    def select_mail(self, mail_index):
        """选择邮件"""
        if self.imap is not None:
            self.imap.select(readonly=False)
            typ1, data1 = self.imap.fetch(mail_index, '(RFC822)')  # 读取邮件
            if typ1 == 'OK':
                try:
                    self.output_mail_text(data1, mail_index)
                except:
                    return 1
                return 0

    def check_new_mail(self):
        """检查是否有新邮件"""
        self.mail_box = mail_box
        if self.imap is not None:
            self.imap.select(readonly=False)
            typy, datay = self.imap.search(None, 'ALL')  # 选择全部邮件
            if typy == 'OK':
                last_new_Email_num = datay[0].split()
                tgprint("新邮件推送服务已开启, 当有新邮件时会通知您")
                while True:
                    time.sleep(delay_time)
                    typ, data = self.imap.search(None, 'ALL')
                    if typ == 'OK':
                        tt = data[0].split()
                        if last_new_Email_num != tt:
                            last_new_Email_num = tt
                            typ1, data1 = self.imap.fetch(tt[-1], '(RFC822)')  # 读取邮件
                            if typ1 == 'OK':
                                global is_mail_deleted
                                if is_mail_deleted:
                                    is_mail_deleted = False
                                    continue
                                self.output_mail_text(data1, tt[-1])
                    else:
                        tgprint("邮件服务器连接异常，正在重启 错误代码: 0x01")
                        restartAuto()
            else:
                tgprint("邮件服务器连接异常，正在重启 错误代码: 0x02")
                restartAuto()
        else:
            tgprint("邮件服务器连接异常，正在重启 错误代码: 0x03")
            restartAuto()

    def output_mail_text(self, data, num):
        temp_text = ''
        msg = email.message_from_string(data[0][1].decode("utf-8"))
        # 获取邮件编码
        msgCharset = email.header.decode_header(msg.get('Subject'))[0][1]
        recv_date = email.header.decode_header(msg.get('Date'))[0][0].replace('+00:00', '')
        # 处理内容编码
        try:
            mail_from = email.header.decode_header(msg.get('From'))[0][0].decode(
                msgCharset)
        except:
            try:
                mail_from = email.header.decode_header(msg.get('From'))[0][0].decode(
                    "unicode_escape")
            except:
                mail_from = email.header.decode_header(msg.get('From'))[0][0]
        try:
            mail_to = email.header.decode_header(msg.get('To'))[0][0].decode(
                msgCharset)
        except:
            try:
                mail_to = email.header.decode_header(msg.get('To'))[0][0].decode(
                    "unicode_escape")
            except:
                mail_to = email.header.decode_header(msg.get('To'))[0][0]
        try:
            subject = email.header.decode_header(msg.get('Subject'))[0][0].decode(
                msgCharset)
        except:
            try:
                subject = email.header.decode_header(msg.get('Subject'))[0][0].decode(
                    "unicode_escape")
            except:
                subject = email.header.decode_header(msg.get('Subject'))[0][0]
        temp_text += "邮件标题: %s\n" % subject
        temp_text += '发件人: ' + mail_from + '\n收件人: ' + mail_to + '\n接收时间: ' + str(
            recv_date) + '\n邮件编号: /*' + str(num).replace('b', '').replace('\'', '') + '*/\n\n'
        temp_header_text = temp_text
        for part in msg.walk():
            if not part.is_multipart():
                name = part.get_param("name")
                if not name:  # 检查是否为附件
                    try:
                        msg_decode = part.get_payload(decode=True).decode(msgCharset)
                    except:
                        try:
                            msg_decode = part.get_payload(decode=True).decode(
                                "unicode_escape")
                        except:
                            msg_decode = part.get_payload(decode=True)
                    if msg_decode.find('<html') == -1 and msg_decode.find('<meta') == -1 and msg_decode.find('<div'):
                        temp_text += msg_decode + '\n'
                    else:
                        temp_text += "<此处出现html文本>\n"
        try:
            tgprint(temp_text)
        except:
            tgprint(temp_header_text + "\n\n消息发送错误，请检查邮件是否过长")

    def delete_mail(self, mail_index):
        """删除邮件"""
        if self.imap is not None:
            self.imap.select(readonly=False)
            self.imap.store(mail_index.encode(), '+FLAGS', '(\\Deleted)')
            self.imap.expunge()
            global is_mail_deleted
            is_mail_deleted = True


def start(update: Update, context):
    if authForUser(update):
        return
    update.message.reply_text(text=start_text)


def info(update: Update, context):
    if authForUser(update):
        return
    text = f"""
    机器人详细配置: 
    
IMAP 服务器地址: `{imap_host}`
SMTP 服务器地址: `{smtp_host}`
邮箱用户名: {username}
密码或授权码: `{password}`
邮箱名: `{mail_box}`
检查新邮件延时: {delay_time}
机器人指定用户ID: `{tg_chat_id}`
机器人 Token: `{tg_bot_token}`
使用代理: {'`' + proxy_url.replace('socks5h', 'socks5') + '`' if proxy_url != '' else '不使用代理'}
    """
    update.message.reply_text(text=text, parse_mode='Markdown')


def getmail(update: Update, context):
    if authForUser(update):
        return
    try:
        mail_len = update.message['text'].split()[1]
    except:
        e2tUtil.get_mail(1)
        tgprint("已为您发送最新一条邮件.\n如想查询更多邮件, 请在命令后输入要输出的邮件个数, 例如: /getmail 3")
        return
    if int(mail_len) <= 0:
        tgprint("无效的输入.")
        return
    e2tUtil.get_mail(mail_len)



def sendmail(update: Update, context) -> int:
    if authForUser(update):
        return ConversationHandler.END
    update.message.reply_text("请输入收件人")
    return 1


def sendmail1(update: Update, context) -> int:
    global receive_people
    receive_people = update.message.text
    update.message.reply_text("请输入邮件主题(标题)")
    return 2


def sendmail2(update: Update, context) -> int:
    global header_text
    header_text = update.message.text
    update.message.reply_text("请输入正文(支持html)")
    return 3


def sendmail3(update: Update, context) -> int:
    global body_text
    body_text = update.message.text
    e2tUtil.send_http_mail(receive_people, header_text, body_text)
    return ConversationHandler.END


def sendmutimail(update: Update, context) -> int:
    if authForUser(update):
        return ConversationHandler.END
    update.message.reply_text("请输入收件人")
    return 1


def sendmutimail1(update: Update, context) -> int:
    global receive_people
    receive_people = update.message.text
    update.message.reply_text("请输入邮件主题(标题)")
    return 2


def sendmutimail2(update: Update, context) -> int:
    global header_text
    header_text = update.message.text
    update.message.reply_text("请输入正文(支持html)")
    return 3


def sendmutimail3(update: Update, context) -> int:
    global body_text
    body_text = update.message.text
    update.message.reply_text("请发送需要添加的附件(目前仅支持一个附件)")
    return 4


def sendmutimail4(update: Update, context) -> int:
    with open(update.message.document.file_name, 'wb') as f:
        update.message.bot.get_file(update.message.document).download(out=f)
    e2tUtil.send_multipart_mail(receive_people, header_text, body_text, update.message.document.file_name)
    return ConversationHandler.END


def deletemail(update: Update, context) -> int:
    if authForUser(update):
        return ConversationHandler.END
    update.message.reply_text(text="请输入要删除的邮件编号")
    return 1


def deletemail1(update: Update, context) -> int:
    global delete_mail_temp
    try:
        delete_mail_temp = str(int(update.message.text))
    except:
        tgprint("输入的值不为数字, 请重新输入正确的邮件编号")
        return 1
    if e2tUtil.select_mail(delete_mail_temp.encode()) == 1 :
        tgprint("输入的索引值无法索引到对应邮件, 请重新输入正确的邮件编号")
        return 1
    update.message.reply_text("您确定要删除这个邮件吗？\n输入 /yes 确定删除, 输入其他任意内容取消删除.")
    return 2


def deletemail2(update: Update, context) -> int:
    e2tUtil.delete_mail(delete_mail_temp)
    update.message.reply_text("删除成功")
    return ConversationHandler.END


def cancel(update: Update, context) -> int:
    update.message.reply_text('您取消了操作')
    return ConversationHandler.END


def restart(update: Update, context):
    if authForUser(update):
        return
    restartAuto()

def restartAuto():
    global e2tUtil
    e2tUtil = Email2TGUtil(imap_host=imap_host, imap_enable_ssl=True, smtp_host=smtp_host, smtp_enable_ssl=True)
    e2tUtil.login(username=username, password=password)
    e2tUtil.configure(delay_time=delay_time, tg_chat_id=tg_chat_id, tg_bot_token=tg_bot_token)
    tgprint("Email2TG重启成功")
    Thread(target=e2tUtil.check_new_mail).start()
    dispatcher.bot.send_message(chat_id=tg_chat_id, text=start_text)

def replymail(update: Update, context):
    if authForUser(update):
        return
    text = """

    """
    update.message.reply_text(text=text)

def help(update: Update, context):
    if authForUser(update):
        return
    update.message.reply_text(text=help_text)


def tgprint(msg):
    print(msg)
    dispatcher.bot.send_message(chat_id=tg_chat_id, text=msg)


def tgmdprint(msg):
    print(msg)
    dispatcher.bot.send_message(chat_id=tg_chat_id, text=msg, parse_mode='Markdown')


def tghtmmlprint(msg):
    print(msg)
    dispatcher.bot.send_message(chat_id=tg_chat_id, text=msg, parse_mode='HTML')


def authForUser(update):
    if str(update.message['chat']['id']) != tg_chat_id:
        update.message.reply_text(text="你未通过验证，无法使用本机器人!")
        tgmdprint(f"""
未验证的用户在未验证的对话中发送了信息.

chat-title: `{update.message['chat']['aaa']}`
chat-type: {update.message['chat']['type']}
chat-id: `{update.message['chat']['id']}`
username: `{update.message['from_user']['username']}`
user-id: `{update.message['from_user']['id']}`
date: {update.message['date']}
text: `{update.message['text']}`

请管理员悉知！
        """)
        return True
    return False


if __name__ == '__main__':

    # 读取配置文件
    config_file = open('test_config.yaml','r',encoding='utf-8').read()
    config = yaml.safe_load(config_file)
    imap_host = config['imap_host']
    smtp_host = config['smtp_host']
    username = config['username']
    password = config['password']
    mail_box = config['mail_box']
    delay_time = config['delay_time']
    tg_chat_id = config['tg_chat_id']
    tg_bot_token = config['tg_bot_token']
    proxy_url = config['proxy_url']

    print("初始化中")
    # 初始化参数
    start_text = """
欢迎使用 Email2TG 机器人!
请发送 /help 来获取详细命令列表.
        """
    help_text = """
欢迎使用 Email2TG 机器人!

以下是详细命令列表: 
1. /start            开始
2. /info             机器人配置详细信息
3. /getmail          获取最近的邮件
4. /sendmail         发送邮件
5. /sendmutimail     发送带附件的邮件
6. /deletemail       删除指定邮件 
7. /replymail        回复邮件
8. /restart          重启 Email2TG
9. /cancel           取消当前操作
        """
    receive_people, header_text, body_text = ('', '', '')
    delete_mail_temp = ''
    is_mail_deleted = False
    if proxy_url == '':
        updater = Updater(token=tg_bot_token)
    else:
        proxy_url = proxy_url.replace('socks5', 'socks5h')
        updater = Updater(token=tg_bot_token, request_kwargs={'proxy_url': proxy_url})
    dispatcher = updater.dispatcher
    dispatcher.add_handler(CommandHandler('start', start))
    dispatcher.add_handler(CommandHandler('info', info))
    dispatcher.add_handler(CommandHandler('getmail', getmail))
    dispatcher.add_handler(ConversationHandler(
        entry_points=[CommandHandler('sendmail', sendmail)],
        states={
            1: [CommandHandler('cancel', cancel), MessageHandler(Filters.text, sendmail1)],
            2: [CommandHandler('cancel', cancel), MessageHandler(Filters.text, sendmail2)],
            3: [CommandHandler('cancel', cancel), MessageHandler(Filters.text, sendmail3)],
        },
        fallbacks=[CommandHandler('cancel', cancel)],
    ))
    dispatcher.add_handler(ConversationHandler(
        entry_points=[CommandHandler('sendmutimail', sendmutimail)],
        states={
            1: [CommandHandler('cancel', cancel), MessageHandler(Filters.text, sendmutimail1)],
            2: [CommandHandler('cancel', cancel), MessageHandler(Filters.text, sendmutimail2)],
            3: [CommandHandler('cancel', cancel), MessageHandler(Filters.text, sendmutimail3)],
            4: [CommandHandler('cancel', cancel), MessageHandler(Filters.document, sendmutimail4)],
        },
        fallbacks=[CommandHandler('cancel', cancel)],
    ))
    dispatcher.add_handler(ConversationHandler(
        entry_points=[CommandHandler('deletemail', deletemail)],
        states={
            1: [CommandHandler('cancel', cancel), MessageHandler(Filters.text, deletemail1)],
            2: [CommandHandler('yes', deletemail2), MessageHandler(Filters.text, cancel)],
        },
        fallbacks=[CommandHandler('cancel', cancel)],
    ))
    dispatcher.add_handler(CommandHandler('replymail', replymail))
    dispatcher.add_handler(CommandHandler('restart', restart))
    dispatcher.add_handler(CommandHandler('help', help))
    # 提交命令
    commands = [
        ("start", "开始"),
        ("info", "机器人配置详细信息"),
        ("getmail", "获取最近的邮件"),
        ("sendmail", "发送邮件"),
        ("sendmutimail", "发送带附件的邮件"),
        ("deletemail", "删除指定邮件"),
        ("replymail", "回复邮件"),
        ("restart", "重启 Email2TG"),
        ("cancel", "取消当前操作"),
        ("help", "帮助列表")
    ]
    if proxy_url != '':
        proxy = telegram.utils.request.Request(proxy_url=proxy_url.replace('socks5h', 'socks5'))
        bot = Bot(token=tg_bot_token, request=proxy)
    else:
        bot = Bot(token=tg_bot_token)
    bot.set_my_commands(commands)
    bot = None
    tgprint("Telegram 服务器已连接")

    e2tUtil = Email2TGUtil(imap_host=imap_host, imap_enable_ssl=True, smtp_host=smtp_host, smtp_enable_ssl=True)
    e2tUtil.login(username=username, password=password)
    e2tUtil.configure(delay_time=delay_time, tg_chat_id=tg_chat_id, tg_bot_token=tg_bot_token)
    tgprint("Email2TG启动成功")
    Thread(target=e2tUtil.check_new_mail).start()
    dispatcher.bot.send_message(chat_id=tg_chat_id, text=start_text)
    updater.start_polling()
