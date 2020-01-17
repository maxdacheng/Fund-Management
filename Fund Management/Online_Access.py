'''
The module for Internet-related content
'''

from bs4 import BeautifulSoup;
import bs4;
import datetime;
from selenium import webdriver;
from selenium.webdriver.chrome.options import Options
import time;
import Global;
from Basic_Class import asset;
from Basic_Class import portfolio;
from Error_Warning import Warning;
import Functions.Web as WebF;
import Functions.Other as OtherF;



def log_in():
    '''Log into the account.

    Input: null
    Output: the BeautifulSoup object of the holding page
    '''

    print("Initialize browser.");
    config=Global.get_value('config');
    option = Options();
    option.add_argument('--start-maximized')
    option.add_argument('--headless')
    option.add_argument('--incognito')
    option.add_argument('--disable-gpu')
    path=config['Path']['ChromeDriverPath'];
    driver = webdriver.Chrome(executable_path=path);
    driver.minimize_window();
    while True:
        print('username:(type in exit to return)');
        username=input();
        if username=='exit':
            driver.close()
            return 0;
        print('password:(type in exit to return)');
        password=OtherF.pwd_input();
        if password=='exit':
            driver.close();
            return 0;
        print();
        print("Begin logging in.");

        try:
            driver.get('https://login.1234567.com.cn/login') 
            time.sleep(1)
            driver.find_element_by_id('tbname').click()
            driver.find_element_by_id('tbname').send_keys(username) 
            driver.find_element_by_id('tbpwd').click()
            driver.find_element_by_id('tbpwd').send_keys(password)
            driver.find_element_by_class_name('submit').click() 
            time.sleep(10)
        except:
            driver.close();
            Warning.warning("Internet problem!");
            return 0;
        try:
            driver.find_element_by_id('myassets_hold').click() 
            time.sleep(3)
            content = driver.page_source.encode('utf-8')
            driver.close()
            break;
        except:
            Warning.warning("Incorrect keywords!");
            continue;
    print("Successful.");
    return content;

def update_asset(a:asset):
    '''
    Update one asset.

    Input: the asset object, the row ID.
    Output: null.'''

    #Form its url and send online request.
    url0 = "http://fund.eastmoney.com/";
    url2=".html";
    url1 = a.ID;
    url = url0 + url1+url2;
    soup=WebF.build_access(url);

    #Error.
    if(soup == 0):
        return;

    #Search on the website.
    if a.type==2:
        a.name,a.curgrowth,a.curneat=WebF.get_monetary(soup);

    if a.type in [0,1]:
        a.curgrowth,a.curneat=WebF.get_bondstock(soup);

def build_portfolio_onlinepart(p:portfolio,content):
    '''
    Build the real-time holding portfolio's content with my account's holding information
    '''
    p.release();

    #Synchroize the bond and stock assets.
    print("Begin synchronizing holding information.");
    soup = BeautifulSoup(content, 'html.parser');
    children = soup.find('table',{'class':'table-hold'}).tbody.children;
    for c in children:
        a=asset(0);
        for d in c.descendants:
            if isinstance(d,bs4.element.Tag):
                if d.get('class')==['lk']:
                    s=''.join(d.string);
                    s1,s2=s.split('（');
                    a.name=s1;
                    a.ID=s2.strip('）');
                if d.get('class')==['info-noborder']:
                    s=''.join(d.string);
                    s1,s2=s.split('（');
                    a.curneat=float(s1[5:]);
                if d.get('class')==['info','info-nopl']:
                    s=''.join(d.string);
                    if s in ['债券型','QDII']:
                        a.type=0;
                    else:
                        a.type=1;
                if d.get('class')==['tor', 'f16', 'desc']:
                    try:
                        s=d.get_text();
                        l=[];
                        for w in s:
                            if OtherF.is_chinese(w):
                                break;
                            else:
                                l.append(w);
                        s=''.join(l);
                        a.curvalue=float(''.join(s.split(',')));
                    except:
                        a.curvalue=0;

                if d.get('class') in [['green-l', 'f16'],['red-l', 'f16']]:
                    try:
                        s=''.join(d.string);
                        a.arevenue=float(''.join(s.split(',')));
                    except:
                        a.arevenue=0;

                if d.get('class') in [['green-l', 'f12'],['red-l', 'f12']]:
                    try:
                        s=''.join(d.string);
                        a.agrowth=float(s.strip("%"))/100;
                    except:
                        a.agrowth=0;
        a.inivalue=a.curvalue-a.arevenue;
        p.content.append(a);

    print("Successful.");


def update_portfolio_onlinepart(p:portfolio):
    '''Update the real-time holding portfolio's growth with online information
    '''
    print("Begin updating up-to-date information.");

       
    #Get current date and time.
    date=datetime.datetime.now().strftime("%Y/%m/%d");
    t= datetime.datetime.now().strftime("%H:%M");
    p.date=str(date);
    p.time=str(t);

    #Form the update list.
    updatelist=p.content.copy(); 
    p.content.clear();
    total=len(updatelist);
    finish=0.0;


    
      



    #Update the assets.
    while len(updatelist)!=0:
        updatelistcopy=updatelist.copy();
        for a in updatelistcopy:
            row=a.row;
            if a.type==1 or a.type==0:
                update_asset(a);
                if a.isinvalid():
                    progress='{0}% finished'.format(int(finish/total*100));
                    print(progress);
                    continue;
                #if a.type==0:
                #    p.bondvalue+=a.curvalue;
                #elif a.type==1:
                #    p.stockvalue+=a.curvalue;

            #Update the monetary funds and save them in the portfolio.
            if a.type==2:
                update_asset(a);
                if a.isempty():
                    progress='{0}% finished'.format(int(finish/total*100));
                    print(progress);
                    continue;
                

            p.content.append(a);
            updatelist.remove(a);
            finish+=1;
            progress='{0}% finished'.format(int(finish/total*100));
            print(progress);

        if len(updatelist)==0:
            break;

        s='Fund ';
        for a in updatelist:
            s+='{0}'.format(a.ID);
            s+=',';
        s=s.rstrip(',');
        s+=' failed in update. Would you try to update them again?';
        if Warning.yes_and_no_warning(s):
            continue;
        else:
            break;

    for a in updatelist:
        p.content.append(a);

    


                
