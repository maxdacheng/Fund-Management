import msvcrt,sys;
import datetime;
import pandas as pd;
import numpy as np;

def is_chinese(uchar):
    if uchar >= u'\u4e00' and uchar<=u'\u9fa5':
        return True
    else:
        return False

def pwd_input():    
    chars = []   
    while True:  
        try:  
            newChar = msvcrt.getch().decode(encoding="utf-8")  
        except: 
            print("Warning: Keywords may not be hidden.") ;
            return input()  
        if newChar in '\r\n':             
             break   
        elif newChar == '\b':  
             if chars:    
                 del chars[-1]   
                 msvcrt.putch('\b'.encode(encoding='utf-8'))   
                 msvcrt.putch( ' '.encode(encoding='utf-8'))  
                 msvcrt.putch('\b'.encode(encoding='utf-8'))                  
        else:  
            chars.append(newChar)  
            msvcrt.putch('*'.encode(encoding='utf-8'))  
    return (''.join(chars) )  


def date64_to_date(d):
    if not pd.isna(d):
        td=datetime.datetime.fromtimestamp((d-np.datetime64('1970-01-01T00:00:00Z')) / np.timedelta64(1, 's'));
        return td.date();
    else:
        return '';


def add_zero_ahead(n):
    if not pd.isna(n):
        s=str(int(n));
        while len(s)<6:
            s='0'+s;
        return s;
    else:
        return '';