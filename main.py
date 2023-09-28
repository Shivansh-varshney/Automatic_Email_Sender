import os
import csv
import time
import smtplib
import pandas as pd
from datetime import date
from plyer import notification
from prettytable import PrettyTable
from email.message import EmailMessage

class autoemail():
    
    def __init__(self, n):
        
        self.n = n
        
        if self.n == 0:
            self.single()
        elif self.n == 1:
            self.many_files()

    def change_dir(self,way):
        os.chdir(way)

    def credentials(self,gmail_id, gmail_password):
        gmail_id = input("Enter your email id please: ")
        gmail_password = input("Enter your password please: ")
        return gmail_id, gmail_password

    def email_content(self,name, body_content):
        email_body_content = body_content.format(name)   
        return email_body_content

    def sendemail(self,msg,GMAIL_ID, GMAIL_PSWD, index, name, to):

        try:
            msg['to'] = to
            s = smtplib.SMTP_SSL('smtp.gmail.com', 465)
            s.login(GMAIL_ID, GMAIL_PSWD)
            s.send_message(msg)
            s.quit()
            del msg["to"]
            
            if (index+1) < 10:
                print(f"{index+1}  | Email sent to {name}.")
            elif 10 <= (index+1) and (index+1) <100:
                print(f"{index+1} | Email sent to {name}.")
            else:
                print(f"{index+1}| Email sent to {name}.")
            return 1,None
        
        except Exception as e:
            print(f"Email could not be sent to {to} due to: ",e)
            return -1,e
        
    def show_failed_emails(self):

        try:
            with open("Excel files\\failed_emails.csv",'r') as f:
                reader_object = csv.reader(f, delimiter = ",")
                t = PrettyTable()
                num = 0
                for i in reader_object:
                    if num == 0:
                        t.field_names = i
                        num += 1
                    else:
                        if i[1] == str(date.today()):
                            t.add_row(i)
                print(t)

        except EOFError:
            f.close()

        except Exception as e:
            print("Data could not be retrieved from the file due to:", e)

    def failed_emails(self,writeind):

        try:
            for key in writeind:
                with open("Excel files\\failed_emails.csv", 'a', newline = '') as f:
                    writer_object = csv.writer(f, delimiter = ",")
                    expression = [writeind[key][0],date.today(),time.strftime("%H:%M:%S", time.localtime()),writeind[key][1]]
                    writer_object.writerow(expression)
                    f.close()
                    
        except Exception as e:
            print("Data could not be added to the file due to:", e)

    def successful_emails(self,successful):

        try:
            for key in successful:
                with open("Excel files\\successful_emails.csv", 'a', newline = '') as f:
                    writer_object = csv.writer(f, delimiter = ",")
                    expression = [successful[key][0],date.today(),time.strftime("%H:%M:%S", time.localtime()), "Successfully",successful[key][1]]
                    writer_object.writerow(expression)
                    f.flush()
                    f.close()

        except Exception as e:
            print("Data could not be added to the file due to:", e)

    def notify_me(self,ttl, msg, icon):
        notification.notify(
                title = ttl,
                message = msg,
                app_icon = icon,
                timeout = 5
                )

    def attach(self,msg):
        #if we want to attach something
        
        name = input("Enter name of file you want to attach with extension:")
        location = "Attachments\\" + name
        
        with open(location, "rb") as f:         
            file_data = f.read()
            file_name = f.name
        return msg.add_attachment(file_data, maintype = 'application', subtype ='pdf', filename= file_name)

    def ask(self):
        ask = input("Do you want to enter more? (Y/N):")
        if ask in 'Nn':
            return False
        elif ask in 'Yy':
            return True
        else:
            print('Wrong input entered. Please try again.')
            ask()

    writeind = {}

    def main(self,i):
        df = pd.read_excel(i)

        GMAIL_ID, GMAIL_PSWD = self.credentials()
    
        successful = {}

        body_content = """Hello {}!."""
        
        for index, item in df.iterrows():

            msg = EmailMessage()

            msg['subject'] = "your subject here."
            msg['from'] = "your name here."
            
            data = self.email_content(item["Name"], body_content)
            msg.set_content(data)

            self.attach(msg)

            if item["Email"] == '__':
                self.writeind[index] = [item["Name"],"No email found"]
                del msg
                continue
            else:
                sent_status,e = self.sendemail(msg,GMAIL_ID, GMAIL_PSWD, index, name = item['Name'], to = item['Email'],)
                n += 1

            if sent_status == 1:
                successful[index] = [item["Name"], item["Email"]]
                del msg
            elif sent_status == -1:
                self.writeind[index] = [item["Name"],e]
                del msg
            else:
                print("some unexpected exception occured please look into the programm manually.")
        
        self.successful_emails(successful)
        
        if len(self.writeind) > 0:
            self.failed_emails(self.writeind)
            return
        return
    
    def many_files(self,):
        # if we want to send emails from more than 1 file at the same time.
        list_of_excel = []

        flag = True
        while flag == True:
            
            name = input("Enter name of file:")
            list_of_excel.append(name)
            flag = self.ask()

        for i in list_of_excel:
            print("Sending emails to the channels in the file",i)
            self.main("Excel files\\" + i + ".xlsx")

    def single(self):
        name = input("Enter name of the Excel file")
        self.main("Excel files\\" + name + ".xlsx")

if __name__ == "__main__":

    os.chdir("C:\\Users\\hp\\OneDrive\\Desktop\\AUTOEMAIL")
    n = int(input("Enter 0 to send email from one file or 1 to send from more than 1 file:"))
    send = autoemail(n)
    
    ttl = "Task has been completed"
    msg = "Emails have been sent to all. Please look for any failed emails."
    icon = "Icons\\Graphicloads-100-Flat-2-Ok.ico"
    send.notify_me(ttl, msg, icon)
    
    if len(send.writeind) > 0:
        if input("Do you want to see list of people to whom email could not be sent?\n Enter 'Y' to continue:") in 'Yy':
            send.show_failed_emails()