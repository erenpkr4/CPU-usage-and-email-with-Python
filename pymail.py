import win32com.client as win32
import psutil
from configparser import ConfigParser
import logging

logging.basicConfig(level=logging.INFO, filename= "log.log", filemode="w",
                    format="%(asctime)s - %(levelname)s - %(message)s")

logger = logging.getLogger(__name__)

config_file = 'config.ini'
config = ConfigParser()
config.read(config_file)

max_usage = float(config['cpu_usage']['max_usage'])
time_interval = int(config['time']['time_interval'])

olApp = win32.Dispatch('Outlook.Application')
olNS = olApp.GetNameSpace('MAPI')

def createNewMail():
    mailItem = olApp.CreateItem(0)
    mailItem.Subject = 'CPU WARNING!'
    mailItem.BodyFormat = 1
    mailItem.Body = f'CPU usage exceeded %{max_usage}!'
    mailItem.To = 'erenpython@outlook.com'
    mailItem.HTMLBody = f'<h2>Average Usage is %{average}'
    mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('eren.peker@bahcesehir.edu.tr')))
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
