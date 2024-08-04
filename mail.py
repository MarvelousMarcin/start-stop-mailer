import win32com.client
import inquirer
from colorama import Fore, Style
from datetime import datetime
import json 
print(Fore.MAGENTA + 'T-MOBILE - START/STOP Mail 😋')

date = datetime.today().strftime('%d.%m.%Y')

class Config: 
    to: str
    title: str
    cc: list[str]
    focus: bool
    
    def __init__(self, to, title, cc, focus):
        self.to = to
        self.title = title
        self.cc = cc
        self.focus = focus

try:
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)

    with open('config.json', 'r') as f:
        data = json.load(f)
        config = Config(data["to"], data["title"], data["cc"], data["focus"])

    answer = inquirer.prompt([inquirer.List('type',message="Mail Type?",choices=['START', 'STOP'],)])

    if answer["type"] == "START":
        title = config.title + " " +  date + " start"
    else:
        title= config.title + " " + date + " stop"
        time = inquirer.prompt(questions = [inquirer.Text('hour', message="Hours(h)")])
        done_thing = inquirer.prompt(questions = [inquirer.Text('done', message="Things done")])
        mail.Body = "1.\t\t" + done_thing["done"] + "\t\t\t" + time["hour"] +"h"
          
    mail.Subject = title
    mail.CC = ";".join(config.cc)
    mail.To = config.to
    
    if config.focus:
        outlook.ActiveWindow().Activate()

    mail.Save()
    print(Fore.MAGENTA + '🎉Mail Created🎉')
    print(Style.RESET_ALL)
except:
    print(Fore.MAGENTA + '🍆There was some problem🍆')
    print(Style.RESET_ALL)