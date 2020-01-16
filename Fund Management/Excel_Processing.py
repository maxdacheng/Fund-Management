import Global;
import xlwings as xl;
import time;
import datetime;
import os;
from Basic_Class import asset;
from Basic_Class import portfolio;
import Functions.Excel as ExcelF;
from Functions.Excel import close_excel;




def init():
    print('Preparing...');  
    
    #Read in initial excel parameters.  
    config=Global.get_value("config");
    global stock_start_row,monetary_start_row,save_start_row,date_start_row;
    stock_start_row=int(config['Excel']['StockStartRow']);
    monetary_start_row=int(config['Excel']['MonetaryStartRow']);
    save_start_row=int(config['Excel']['SaveStartRow']);
    date_start_row=int(config['Excel']['DateStartRow']);
    path=config["Path"]["ExcelPath"];

    #Open the excel file.
    os.startfile(path);
    time.sleep(5);
    sheet0,sheet1,sheet2=ExcelF.load_excel(); 
    Global.set_value('sheet0',sheet0);
    Global.set_value('sheet1',sheet1);
    Global.set_value('sheet2',sheet2);



def count():
    sheet1=Global.get_value('sheet1');
    sheet2=Global.get_value('sheet2');
    a='a{0}'.format(stock_start_row-1)
    b='a{0}'.format(monetary_start_row-1)
    c='a{0}'.format(save_start_row-1)
    obj=zip([sheet1,sheet1,sheet1,sheet2],[a,b,c,'f1']);
    rowcount=tuple(map(ExcelF.count_rows,obj));
    Global.set_value('rowcount',rowcount);


def read_asset_in_holding_basic(a:asset):
    row=a.row;
    sheet1=Global.get_value('sheet1');
    if a.type==0 or a.type==1:
        a.ID=sheet1.range('a{0}'.format(row)).value;
        a.name=sheet1.range('c{0}'.format(row)).value;                
        a.inivalue=float(sheet1.range('d{0}'.format(row)).value);
        a.curvalue=float(sheet1.range('e{0}'.format(row)).value);
        a.arevenue=float(sheet1.range('j{0}'.format(row)).value);
        a.agrowth=a.arevenue/a.inivalue;

    if a.type==2:
        a.ID=sheet1.range('a{0}'.format(row)).value;
        a.inivalue=float(sheet1.range('d{0}'.format(row)).value);

    if a.type==3:
        a.ID=sheet1.range('a{0}'.format(row)).value;
        a.name=sheet1.range('c{0}'.format(row)).value; 
        a.inivalue=float(sheet1.range('d{0}'.format(row)).value);
        a.curgrowth=float(sheet1.range('h{0}'.format(row)).value);


def read_asset_in_holding_advance(a:asset):
    row=a.row;
    sheet1=Global.get_value('sheet1');
    if a.type==0 or a.type==1:                
        a.curneat=float(sheet1.range('g{0}'.format(row)).value);                  
        a.curgrowth=float(sheet1.range('h{0}'.format(row)).value); 

    if a.type==2:
        a.name=sheet1.range('c{0}'.format(row)).value;   
        a.curneat=float(sheet1.range('h{0}'.format(row)).value);
        a.curgrowth=float(sheet1.range('i{0}'.format(row)).value);
        a.curvalue=float(sheet1.range('f{0}'.format(row)).value);
        a.arevenue=a.curvalue-a.inivalue;
        a.agrowth=a.arevenue/a.inivalue;

    if a.type==3:
        a.curvalue=float(sheet1.range('f{0}'.format(row)).value);
        a.arevenue=a.curvalue-a.inivalue;
        a.agrowth=a.arevenue/a.inivalue;


def read_asset_in_holding(a:asset,mode):
    row=a.row;
    sheet1=Global.get_value('sheet1');

    #Decide its type.
    type=['债券型','股票型','货币型'];
    if sheet1.range('b{0}'.format(row)).value not in type:
        a.type=3;
    else:
        a.type=type.index(sheet1.range('b{0}'.format(row)).value);

    if mode=='all':
        read_asset_in_holding_basic(a);
        read_asset_in_holding_advance(a);

    if mode=='basic':
        read_asset_in_holding_basic(a);

    if mode=='advance':
        read_asset_in_holding_advance(a);


def read_portfolio_in_holding(p,mode):
    #Fill in the portfolio with the data in the table without update

    sheet0=Global.get_value('sheet0');
    sheet1=Global.get_value('sheet1');
    rowcount=Global.get_value('rowcount');

    if mode=='all':
        #Get current date and time.
        date=datetime.datetime.now().strftime("%Y/%m/%d");
        t= datetime.datetime.now().strftime("%H:%M");
        p.date=str(date);
        p.time=str(t);

        p.inivalue=float(sheet0.range('b3').value);
        p.sparevalue=float(sheet0.range('h8').value);
        p.pvalue+=p.sparevalue;

        for row in range(stock_start_row,stock_start_row+rowcount[0]-1):
            a=asset(0);
            a.row=row;
            read_asset_in_holding(a,mode='all');
            if a.type==0:
                p.bondvalue+=a.curvalue;
            else:
                p.stockvalue+=a.curvalue;
            p.pvalue+=a.curvalue;
            p.content.append(a);
        
        for row in range(monetary_start_row,rowcount[1]+monetary_start_row-1):
            a=asset(2);
            a.row=row;
            read_asset_in_holding(a,mode='all');
            p.monevalue+=a.curvalue;
            p.pvalue+=a.curvalue;
            p.content.append(a);

        for row in range(save_start_row,save_start_row+rowcount[2]-1):
            a=asset(3);
            a.row=row;
            read_asset_in_holding(a,mode='all');
            p.savevalue+=a.curvalue;
            p.pvalue+=a.curvalue;
            p.content.append(a);

        for a in p.content:
            a.ratio=a.curvalue/p.pvalue;

        p.prevenue=p.pvalue-p.inivalue;
        p.pgrowth=p.prevenue/p.inivalue;

    if mode=='basic':
        for row in range(stock_start_row,rowcount[0]+stock_start_row-1):
            a=asset(0);
            a.row=row;
            read_asset_in_holding(a,mode='basic');
            p.content.append(a);
        
        for row in range(monetary_start_row,rowcount[1]+monetary_start_row-1):
            a=asset(2);
            a.row=row;
            read_asset_in_holding(a,mode='basic');
            p.content.append(a);

        for row in range(save_start_row,rowcount[2]+save_start_row-1):
            a=asset(3);
            a.row=row;
            read_asset_in_holding(a,mode='basic');
            p.content.append(a);


def read_portfolio_in_history(begin:int,end:int):
    '''Read in one historical record.
    Input: the beginning and the end.
    Output: the portfolio object.'''

    assert begin<=end;

    p=portfolio();
    sheet2=Global.get_value('sheet2');
    p.release();
    count=end-begin+1;

    #Read in the date and time.
    p.date=sheet2.range('b{0}'.format(begin)).value;
    p.time=sheet2.range('c{0}'.format(begin)).value;
    p.pvalue=sheet2.range('l{0}'.format(begin)).value;
    p.prevenue=sheet2.range('m{0}'.format(begin)).value;
    p.pgrowth=sheet2.range('n{0}'.format(begin)).value;

    #Read in each of its assets.
    for i in range(begin,end+1):
        if sheet2.range('e{0}'.format(i)).value == "债券型":
            a=asset(0);
        elif sheet2.range('e{0}'.format(i)).value == "股票型":
            a=asset(1);
        elif sheet2.range('e{0}'.format(i)).value == "货币型":
            a=asset(2);
        elif sheet2.range('e{0}'.format(i)).value == "定期":
            a=asset(3);
        a.ID=sheet2.range('d{0}'.format(i)).value;
        a.name=sheet2.range('f{0}'.format(i)).value;
        a.curvalue=float(sheet2.range('g{0}'.format(i)).value);
        a.curneat=float(sheet2.range('i{0}'.format(i)).value);
        a.arevenue=float(sheet2.range('j{0}'.format(i)).value);
        a.agrowth=float(sheet2.range('k{0}'.format(i)).value);
        p.content.append(a);
    return p;


def write_portfolio_in_holding(p:portfolio):
    #Show the data in up-to-date position sheet.
    #Input: the portfolio object.
    #Output: null.

    sheet1=Global.get_value('sheet1');

    #Show the general data.
    sheet1.range('c{0}'.format(date_start_row)).value=p.date;
    sheet1.range('d{0}'.format(date_start_row)).value=p.time;


    for a in p.content:
        row=a.row;
        if a.type==1 or a.type==0:
            sheet1.range('a{0}'.format(row)).value=a.ID;
            if a.type==0:
                sheet1.range('b{0}'.format(row)).value="债券型";
            else:
                sheet1.range('b{0}'.format(row)).value="股票型";
            sheet1.range('c{0}'.format(row)).value=a.name;
            sheet1.range('d{0}'.format(row)).value=a.inivalue;
            sheet1.range('e{0}'.format(row)).value=a.curvalue; 
            sheet1.range('f{0}'.format(row)).value=a.ratio;
            sheet1.range('g{0}'.format(row)).value=a.curneat;   
            sheet1.range('h{0}'.format(row)).value=a.curgrowth;            
            sheet1.range('j{0}'.format(row)).value=a.arevenue;                  
            sheet1.range('k{0}'.format(row)).value=a.agrowth;   
        if a.type==2:
            sheet1.range('c{0}'.format(row)).value=a.name;
            sheet1.range('h{0}'.format(row)).value=a.curneat;
            sheet1.range('i{0}'.format(row)).value=a.curgrowth;
            sheet1.range('g{0}'.format(row)).value=a.ratio;
            sheet1.range('f{0}'.format(row)).value=a.curvalue;
            sheet1.range('k{0}'.format(row)).value=a.arevenue;
            sheet1.range('l{0}'.format(row)).value=a.agrowth;
        if a.type==3:
            sheet1.range('f{0}'.format(row)).value=a.curvalue;
            sheet1.range('g{0}'.format(row)).value=a.ratio;
            sheet1.range('i{0}'.format(row)).value=a.arevenue;
            sheet1.range('j{0}'.format(row)).value=a.agrowth;


def write_portfolio_in_history(p:portfolio):
    #Show the data in historical record sheet.
    #Input: the portfolio object.
    #Output: null.

    sheet2=Global.get_value('sheet2');
    rowcount=Global.get_value('rowcount');

    #Form its ID.
    i=rowcount[3];
    while sheet2.range('a{0}'.format(i)).value == None:
        i=i-1;
    try:
        lastid=int(sheet2.range('a{0}'.format(i)).value);
    except:
        print("Warning: Unexpected modification to the ID column.");
        

    #Show its ID, data, time, portfolio value, revenue and growth rate.
    count=rowcount[3]+1;
    sheet2.range('a{0}'.format(count)).value=str(lastid+1);
    sheet2.range('b{0}'.format(count)).value=p.date;
    sheet2.range('c{0}'.format(count)).value=p.time;
    sheet2.range('l{0}'.format(count)).value=p.pvalue;
    sheet2.range('m{0}'.format(count)).value=p.prevenue;
    sheet2.range('n{0}'.format(count)).value=p.pgrowth;

    #Show the detailed data of each of its assets.
    for a in p.content:
        if a.type == 1 or a.type==0:
            sheet2.range('d{0}'.format(count)).value=str(a.ID);
            sheet2.range('f{0}'.format(count)).value=str(a.name);
            sheet2.range('g{0}'.format(count)).value=str(a.curvalue);
            sheet2.range('h{0}'.format(count)).value=str(a.ratio);
            sheet2.range('i{0}'.format(count)).value=str(a.curneat);
            sheet2.range('j{0}'.format(count)).value=str(a.arevenue);
            sheet2.range('k{0}'.format(count)).value=str(a.agrowth);
            if a.type == 1:
                sheet2.range('e{0}'.format(count)).value="股票型";
            else:
                sheet2.range('e{0}'.format(count)).value="债券型";

        if a.type == 2:
            sheet2.range('d{0}'.format(count)).value=str(a.ID);
            sheet2.range('f{0}'.format(count)).value=str(a.name);
            sheet2.range('e{0}'.format(count)).value="货币型";
            sheet2.range('g{0}'.format(count)).value=str(a.curvalue);
            sheet2.range('h{0}'.format(count)).value=str(a.ratio);
            sheet2.range('i{0}'.format(count)).value="1";
            sheet2.range('j{0}'.format(count)).value=str(a.arevenue);
            sheet2.range('k{0}'.format(count)).value=str(a.agrowth);

        if a.type==3:
            sheet2.range('f{0}'.format(count)).value=str(a.name);
            sheet2.range('e{0}'.format(count)).value="定期";
            sheet2.range('g{0}'.format(count)).value=str(a.curvalue);
            sheet2.range('h{0}'.format(count)).value=str(a.ratio);
            sheet2.range('i{0}'.format(count)).value="1";
            sheet2.range('j{0}'.format(count)).value=str(a.arevenue);
            sheet2.range('k{0}'.format(count)).value=str(a.agrowth);

        count+=1;


def write_portfolio_in_intro(p:portfolio):
    sheet0=Global.get_value('sheet0');

    sheet0.range('b4').value=p.pvalue;
    sheet0.range('b8').value=p.bondvalue;
    sheet0.range('e8').value=p.monevalue;
    sheet0.range('b9').value=p.stockvalue;
    sheet0.range('e9').value=p.savevalue;


def build_portfolio_excelpart(p:portfolio,content):
    '''Further build the holding portfolio's content based on excel information
    Attention: This must be used after build_portfolio_onlinepart in Online_Access
    '''
    sheet1=Global.get_value('sheet1');
    sheet2=Global.get_value('sheet2');
    rowcount=Global.get_value('rowcount');
    #stock_start_row=Global.get_value('stock_start_row');
    #monetary_start_row=Global.get_value('monetary_start_row');
    #save_start_row=Global.get_value('save_start_row');
    row=stock_start_row;
    for a in p.content:
        a.row=row;
        row+=1;
    for i in range(monetary_start_row,rowcount[1]+monetary_start_row-1):
        a=asset(2);
        a.row=i;
        read_asset_in_holding(a,mode='basic');
        p.content.append(a);
    for i in range(save_start_row,rowcount[2]+save_start_row-1):
        a=asset(3);
        a.row=i;
        read_asset_in_holding(a,mode='basic');
        p.content.append(a);


def update_portfolio_excelpart(p:portfolio):
    '''
    Further build the holding portfolio's content based on excel information
    Attention: This must be used after update_portfolio_onlinepart in Online_Access
    '''
    sheet0=Global.get_value('sheet0');
    sheet1=Global.get_value('sheet1');
    sheet2=Global.get_value('sheet2');
    rowcount=Global.get_value('rowcount');

    #Locate the last record except today's.
    i=rowcount[3];
    while(1):
        lastdate=sheet2.range('b{0}'.format(i)).value;
        if lastdate != None:
            lastDate=lastdate.strftime("%Y/%m/%d");
            if lastDate != p.date:
                break;
        i-=1;
    period=datetime.datetime.now().day-lastdate.day-1;

    #Read in that historical record. This will be used to calculate the earning of monetary funds.
    end=ExcelF.find_the_end(sheet2,'b',i,rowcount[3]);
    tempp=read_portfolio_in_history(i,end);

    for a in p.content:
        if a.type==0:
            p.stockvalue+=a.curvalue;
        elif a.type==1:
            p.bondvalue+=a.curvalue;
        elif a.type==2:
            for b in tempp.content:
                    if a.ID == b.ID:
                        a.curvalue=b.curvalue+(b.curvalue+period*a.curgrowth/360)/10000*a.curneat;
            a.arevenue=a.curvalue-a.inivalue;
            a.agrowth=a.arevenue/a.inivalue;
            p.monevalue+=a.curvalue;
        elif a.type==3:
            startdate=a.ID;
            today=datetime.datetime.now();
            enddate=today;
            for b in tempp.content:
                if a.name==b.name:
                    a.curvalue=b.curvalue+b.curvalue*a.curgrowth/360*(enddate-lastdate).days;
            a.arevenue=a.curvalue-a.inivalue;
            a.agrowth=a.arevenue/a.inivalue;
            p.savevalue+=a.curvalue;
        p.pvalue+=a.curvalue;

    p.inivalue=float(sheet0.range('b3').value);
    p.sparevalue=float(sheet0.range('h8').value);
    p.pvalue+=p.sparevalue;
    p.prevenue=p.pvalue-p.inivalue;
    p.pgrowth=p.prevenue/p.inivalue;

    #Calculate the ratio of each asset.
    for a in p.content:
        a.ratio=a.curvalue/p.pvalue;