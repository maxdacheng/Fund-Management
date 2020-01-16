'''
Main functions of this program

Attention: For all functions here, 'return 0' means failure,  'return 1' means success, and 'return -1' means an unknown command'
'''

import Global;
from Basic_Class import asset,portfolio;
import Excel_Processing as Excel;
import Online_Access as Online;
import Assisting_Tool as Assist;
import Data_Analysis as Data;
import Backup;
from Error_Warning import Warning;



def update(page,p:portfolio):
    '''
    Update the current holding portfolio

    Input: the cache page, the cache portfolio
    Output: result
    '''

    try:
        if page==0:
            s='Keywords are needed to synchronize with the account. Would you need to synchronize?';
            if Warning.yes_and_no_warning(s):
                page=Online.log_in();
                if page!=0:
                    p.release();
                    Online.build_portfolio_onlinepart(p,page);
                    Excel.build_portfolio_excelpart(p,page);
                else:
                    return 0;
        if p.isempty():
            Excel.read_portfolio_in_holding(p,mode='basic');
        Online.update_portfolio_onlinepart(p);
        Excel.update_portfolio_excelpart(p);
        Excel.write_portfolio_in_holding(p);
        Excel.write_portfolio_in_intro(p);
    except:
        return 0;
    else:
        return 1;


def record(p:portfolio):
    '''
    Record the current holding portfolio in the historical record page.

    Input: the cache portfolio
    Output: result
    '''

    try:
        if p.isempty():
            s1='No portfolio in cache.';
            s2='Would you save the data in current holding and record them in history?';
            if Warning.yes_and_no_warning(s1,s2):
                Excel.read_portfolio_in_holding(p,mode='all');
            else:
                return 0;
                         
        Excel.write_portfolio_in_history(p);
    except:
        return 0;
    else:
        return 1;


def updateintro(p:portfolio):
    '''
    Refresh the introduction page.

    Input: the cache portfolio
    Output: result
    '''

    try:
        p.release();
        Excel.read_portfolio_in_holding(p,mode='all');
        Excel.write_portfolio_in_intro(p);
    except:
        return 0;
    else:
        return 1;


def prints(command,p:portfolio):
    '''
    Print something about the current cache portfolio.

    Input: command string
    Output: result
    '''

    try:
        if p.isempty():
            s1='Warning: No portfolio in cache.';
            s2='Would you read the data in current holding?';
            if Warning.yes_and_no_warning(s1,s2):
                Excel.read_portfolio_in_holding(p,mode='all');
            else:
                return 0;
        if Assist.printout(command,p):
            return 1;
        else:
            return -1;
    except:
        return 0;


def calculate(command):
    '''
    Act like a calculator.

    Input: command string
    Output: result
    '''

    try:
        if Assist.calculator(command):
            return 1;
        else:
            return -1;
    except:
        return 0;


def plot(command):
    '''
    Plot the tendency chart.

    Input: command string
    Output: result
    '''
    try:
        if Data.plotting(command):
            return 1;
        else:
            return -1;
    except:
        return 0;


def help(command):
    '''
    Print the helping page.

    Input: command string
    Output: result
    '''
    try:
        if Assist.help(command):
            return 1;
        else:
            return -1;
    except:
        return 0;


def auto_backup():
    '''
    Automatically backup the historical record.

    Input: null
    Output: result
    '''
    try:
        config=Global.get_value('config');
        online_bool=int(config['Backup']['AutoOnlineBackup']);
        local_bool=int(config['Backup']['AutoLocalBackup']);
        if Backup.check():
            if Warning.yes_and_no_warning('You may need an automatic backup of the historical record.','Would you like to conduct it?'):
                pass;
            else:
                return 1;
            if online_bool:
                Backup.online_backup('auto');
            else:
                print('Automatic online backup is off.')
            if local_bool:
                Backup.local_backup('auto');
            else:
                print('Automatic local backup is off.')

            return 1;
    except:
        return 0;


#def manual_backup():
#    '''
#    Manually backup the historical record.

#    Input: null
#    Output: result
#    '''
#    try:
#        Backup.online_backup
#        return 1;
#    except:
#        return 0;




