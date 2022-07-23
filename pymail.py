import win32com.client as win32
import psutil
from configparser import ConfigParser
import logging
import base64

with open('error_image.png', 'rb') as img_file:     # open error image as img_file
    b64_string = base64.b64encode(img_file.read())  # encode it with b64

final_string = b64_string.decode('utf-8')           # decode with utf-8 to remove the b' 


logging.basicConfig(level=logging.INFO, filename= "log.log", filemode="w",  # declare config properties
                    format="%(asctime)s - %(levelname)s - %(message)s")

config_file = 'config.ini'
config = ConfigParser()                     # initialize config file
config.read(config_file)


max_usage = float(config['cpu_usage']['max_usage'])         # read config to assign variables
time_interval = int(config['time']['time_interval'])


olApp = win32.Dispatch('Outlook.Application')               # dispatch
olNS = olApp.GetNameSpace('MAPI')


def createNewMail():
    mailItem = olApp.CreateItem(0)
    mailItem.Subject = 'CPU WARNING!'
    mailItem.BodyFormat = 1
    mailItem.Body = f'CPU usage exceeded %{max_usage}!'
    mailItem.To = 'erenpython@outlook.com'

    attachment_path = "C:/Workspace/Pymail/view.jpg"
    mailItem.Attachments.Add(attachment_path)         # quote this to remove attachment but still show it on html body     
    attachment = mailItem.Attachments.Add(attachment_path)  
    attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MyId1")    # get cid of attachment


    #HTML STRING HERE----------
    html_string = f"""
            <h2>CPU usage exceeded %{max_usage}!</h2>
            <p>Average Usage is %{average}</p>
            <img src="data:image/png;base64,{final_string}"><br>
            <img src="cid:MyId1">
    """
    #---------------------------

    mailItem.HTMLBody = (html_string) # assign HTML body to string

    mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('eren.peker@bahcesehir.edu.tr')))     # declare email sender

    return mailItem


minuteArray = []

while True:
    usage = psutil.cpu_percent(interval = 1)
    print(usage)
    logging.info(f"Usage measured: %{usage}")
    minuteArray.append(usage)
    if len(minuteArray) >= time_interval:
        average = sum(minuteArray)/len(minuteArray)

        if  average >= max_usage:
            print(f'Average Usage Exceeded %{int(max_usage)}!')
            logging.warning(f"Avg. Usage Exceeded %{max_usage}. Measured: %{usage}")
            minuteArray.clear()
            createNewMail().Send()
            break

        else:
            print('Average Usage Is Acceptible')
            logging.info(f"Avg. Is Acceptible, measured: %{usage}")
            minuteArray.clear()
            break

#https://banner2.cleanpng.com/20180219/ddw/kisspng-error-http-404-icon-cent-sign-cliparts-5a8b089c89d296.8666809315190611485645.jpg
#mailItem.HTMLBody = (f'<h2>CPU usage exceeded %{max_usage}!</h2> <p>Average Usage is %{average}</p> <img src="data:image/png;base64,' +final_string + '">')