"""
This is An Automation Project to Send Email To the Recipants on their Birthday 
by taking input through Csv(Excel) sheet.

 

"""
import schedule
from sys import*
import csv
import winsound    #for Alert 
import os
import datetime,time
import smtplib, ssl


#mail Sending Logic
def send_mail(Reciver):

    port = 587
    
    smtp_server = "smtp.gmail.com"
    sender_email = "gaurav.arun.pekhale@gmail.com"
    password = 'Pekhale@123'
    receiver_email = Reciver
    
    fp=open('mailfile.txt','r',encoding='UTF-8')
    Title = fp.readline()
    fp.seek(0,1)
    body=fp.read()
    log=open('logfile.txt','a',encoding='UTF-8')

    try :
        context = ssl.create_default_context()
        server=smtplib.SMTP(smtp_server, port)

        server.starttls(context=context)

        server.login(sender_email, password)
        
        server.sendmail(sender_email,receiver_email, f"Subject: {Title}\n{body}")

        data=" %s   %s\n"%(Reciver,str(datetime.date.today()))
        log.write(data)
        frequency = 2500
        duration = 100
        winsound.Beep(frequency, duration)


    except Exception as E:
        print("Exception Occurred : ",E)
    finally:
        fp.close()
        log.close()
        server.quit()


def Check_Birthdate(paths):
    try:
        isExist = os.path.exists(paths)
        if isExist:
            csvfile=open(paths,'r',newline='')
            reader = csv.reader(csvfile, delimiter=',', quotechar='|')    #return data  in List format

            for row  in reader:
                birth = datetime.datetime.strptime(row[2],"%d-%m-%Y").date()
                today = datetime.date.today()
                if birth.day == today.day and birth.month == today.month:
                    send_mail(row[1])
        else :
            print("Invalide Path")

    except Exception as E:
        print("EXception Occured :",E)


def main():
    print("Application name :",argv[0])
    
   
    if (len(argv)<2):
        print("Error : Invalide Number of Arguments")
        print("Please Give Flag '-u' for Usage and '-h' for Help")
        exit(0)
    if argv[1] == '-h' or argv[1] == '-H':
        print("HELP : This Script designed for Sending the Mail through python ")
        exit(0)
    if argv[1] == '-u' or argv[1] == '-U':
        print("Usage : Application_name  Recipant_Email_Address ")
        exit(0)

    
   #To send mail at sharp 12:00am
    schedule.every().day.at("00:00").do(Check_Birthdate,argv[1])
    while True:
        schedule.run_pending()
        time.sleep(2)

if __name__=="__main__":
    main()
