'''
Backup the historical page if necessary.
'''

import Global;
import pandas as pd;
import numpy as np;
import datetime;
import time;
import os;
import Functions.Web as WebF;
import Functions.Other as OtherF;


def check():
    '''
    Check whether the system needs an auto-backup.

    Input: null
    Output: null
    '''
    config=Global.get_value('config');
    asg_interval=int(config['Backup']['AutoBackupInterval']);

    data=pd.read_excel('fund management.xlsx',sheet_name=2,usecols=[0,1],skiprows=1);
    data=data.dropna(axis=0);
    datelist=data.iloc[:,1].values;
    record_startdate=OtherF.date64_to_date(datelist[0]);
    record_enddate=OtherF.date64_to_date(datelist[-1]);
    cur_interval=(record_enddate-record_startdate).days;

    if asg_interval<=cur_interval:
        return True;
    else:
        return False;


def local_backup(mode):
    '''
    Backup the historical record locally.

    Input: mode string(manual or auto)
    Output: null
    '''
    print('Start local backup.')
    config=Global.get_value('config');
    backup_path=config['Path']['BackupPath'];

    data=pd.read_excel('fund management.xlsx',sheet_name=2,skiprows=1);
    datelist=data.iloc[:,1];   
    datelist=datelist.dropna(axis=0);
    datelist=datelist.values;

    record_startdate=OtherF.date64_to_date(datelist[0]);
    record_enddate=OtherF.date64_to_date(datelist[-1]);
    record_startdate_str=record_startdate.strftime('%Y-%m-%d');
    record_enddate_str=record_enddate.strftime('%Y-%m-%d');

    if mode=='manual':
        filename='ManualBackup({0} to {1}).xlsx'.format(record_startdate_str,record_enddate_str);
    else:
        filename='AutoBackup({0} to {1}).xlsx'.format(record_startdate_str,record_enddate_str);

    data.iloc[:,1]=[OtherF.date64_to_date(d) for d in data.iloc[:,1].values];
    data.iloc[:,3]=[OtherF.add_zero_ahead(d) for d in data.iloc[:,3].values];
    data.to_excel(backup_path+'\\'+filename,index=False);
    print('Backup success.');


def online_backup(mode):
    '''
    Backup the historical record online.

    Input: mode string(manual or auto)
    Output: null
    '''
    print('Start online backup.')
    config=Global.get_value('config');
    address=config['Email']['Address'];
    code=config['Email']['Code'];

    data=pd.read_excel('fund management.xlsx',sheet_name=2,skiprows=1);

    datelist=data.iloc[:,1];
    datelist=datelist.dropna(axis=0);
    datelist=datelist.values;

    record_time=datetime.datetime.now();
    record_startdate=OtherF.date64_to_date(datelist[0]);
    record_enddate=OtherF.date64_to_date(datelist[-1]);
    record_time_str=record_time.strftime('%Y-%m-%d %H:%M:%S');
    record_startdate_str=record_startdate.strftime('%Y-%m-%d');
    record_enddate_str=record_enddate.strftime('%Y-%m-%d');


    if mode=='manual':
        filename='ManualBackup({0} to {1}).xlsx'.format(record_startdate_str,record_enddate_str);
        title_text='ManualBackup({0} to {1})'.format(record_startdate_str,record_enddate_str);
    else:
        filename='AutoBackup ({0} to {1}).xlsx'.format(record_startdate_str,record_enddate_str);
        title_text='AutoBackup ({0} to {1})'.format(record_startdate_str,record_enddate_str);

    email_text_file=open('Texts\\Backup_Email.txt','r');
    email_text=email_text_file.read();
    email_text=email_text.format(record_time_str);

    data.iloc[:,1]=[OtherF.date64_to_date(d) for d in data.iloc[:,1].values];
    data.iloc[:,3]=[OtherF.add_zero_ahead(d) for d in data.iloc[:,3].values];
    data.to_excel(filename,index=False);

    WebF.send_email(username_send=address,username_recv=address,password=code,file=filename,title=title_text,text=email_text);
    time.sleep(5);
    os.remove(filename);
    print('Backup success.');

