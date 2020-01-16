import Global;
from Basic_Class import portfolio;
import matplotlib.pyplot as plt;
import pandas as pd;
import numpy as np;
import datetime;



def plot(x,y,xlabel,ylabel,title):
    fig = plt.figure()
    ax = fig.add_subplot(1, 1, 1);
    ax.plot(x,y,'o-');
    ax.set_xlabel(xlabel);
    ax.set_ylabel(ylabel);
    ax.set_title(title);
    plt.show();
    

def plotting(command):
    rule=Global.get_value('rule');
    comm=command[4:];
    content=comm.strip('() ');
    if content not in rule['plot'].keys():
        return False;
    col_num=rule['plot'][content];
    data=pd.read_excel('fund management.xlsx',sheet_name=2,usecols=[1,col_num],skiprows=1);
    data=data.dropna(axis=0);
    x=np.array(data.iloc[:,0].values);
    y=np.array(data.iloc[:,1].values);
    xlabel='date';
    ylabel=content;
    plot(x,y,xlabel,ylabel,'{0}-{1}'.format(xlabel,ylabel));
    return True;


