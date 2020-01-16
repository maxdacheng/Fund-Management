from bs4 import BeautifulSoup;
import bs4;
import xlwings as xl;
import urllib.request;
import requests;
import urllib;
import smtplib
import base64
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart



def build_access(url:str):
    #Send requests to the corresponding websites.
    #Input: url
    #Output: the BeautifulSoup object.
    try:
        headers = {
            'User-Agent': r'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) '
                            r'Chrome/45.0.2454.85 Safari/537.36 115Browser/6.0.3',
            'Referer': url,
            'Connection': 'keep-alive'
        }
        print("Sending request to ",url);
        req=requests.get(url,headers=headers,timeout=15);
        page=req.text;
        soup = BeautifulSoup(page,'html.parser');
        print("Successful");
        return soup;
    except:
        print("Failing");
        return 0;

def get_monetary(soup):
    name='none';
    annualyield=0.0;
    earning=0.0;
    for link in soup.find_all("span"): 
        #Search for its name.
        if link.get("class")==["funCur-FundName"]:
            name=link.string;
            

        #Get 7-day annual yield.
        if link.get("class") == ["fix_zzl","bold","ui-color-red"]:
            s=link.string;
            a=''.join(s);
            b=a.strip('%+ ');
            try:
                annualyield=float(b)/100.0;
            except:
                annualyield=0;
            
            
                
        #Get the earning per 10000.
        if link.get('class') == ["fix_dwjz","bold","ui-color-red"]:
            s=link.string;         
            a=''.join(s);
            try:
                earning=float(a);
            except:
                earning=0;
            

    return name,annualyield,earning;

def get_bondstock(soup):
    curgrowth=0.0;
    curneat=0.0;
    lastgrowth=0.0;
    lastneat=0.0;
    for link in soup.find_all("span"): 
        #Get today's growth rate
        if link.get("id") == "gz_gszzl":
            s=link.string;
            a=''.join(s);
            b=a.strip('%+ ');
            try:
                curgrowth=float(b)/100.0;
            except:
                curgrowth=0;
            continue;

        #Get today's neat value.
        if link.get("id") == "gz_gsz":
            s=link.string;         
            a=''.join(s);
            try:
                curneat=float(a);
            except:
                curneat=0;
            continue;

        #Get yesterday's growth rate. This is the special consideration for funds related to the US market.
        if link.get("class") == ["fix_zzl","bold","ui-color-green"] or link.get("class") ==["fix_zzl","bold","ui-color-red"]:
            s=link.string;
            a=''.join(s);
            b=a.strip('%+ ');
            try:
                lastgrowth=float(b)/100.0;
            except:
                lastgrowth=0;
            continue;
                
        #Get yesterday's neat value. This is the special consideration for funds related to the US market.
        if link.get("class") == ["fix_dwjz","bold","ui-color-green"] or link.get("class")==["fix_dwjz","bold","ui-color-red"]:
            s=link.string;         
            a=''.join(s);
            try:
                lastneat=float(a);
            except:
                lastneat=0;
            continue;

    #Replace today's data with yesterday's if today's data has not yet been updated online.
    if curgrowth==0.0:
        curgrowth=lastgrowth;
    if curneat==0.0:
        curneat=lastneat;

    return curgrowth,curneat;

def send_email(username_send,username_recv,password,file,title,text):
    print('Sending email.')
    mailserver = "smtp.qq.com"  
    mail = MIMEMultipart()
    att = MIMEText(open(file, 'rb').read(),"base64", "utf-8")
    att["Content-Type"] = 'application/octet-stream'
    new_file='=?utf-8?b?' + base64.b64encode(file.encode()).decode() + '?='
    att["Content-Disposition"] = 'attachment; filename="%s"'%new_file
    mail.attach(att);
    mail.attach(MIMEText(text));
    mail['Subject'] = title;
    mail['From'] = username_send  
    mail['To'] = username_recv  
    smtp=smtplib.SMTP_SSL('smtp.qq.com',port=465) 
    smtp.login(username_send,password) 
    smtp.sendmail(username_send,username_recv,mail.as_string())
    smtp.quit() 
    print ('Success.')
