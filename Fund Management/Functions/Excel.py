import xlwings as xl;

def load_excel():
    excel=xl.Book('Fund Management.xlsx');
    return excel.sheets[0],excel.sheets[1],excel.sheets[2];


def close_excel():
    excel=xl.Book('Fund Management.xlsx');
    excel.save();
    excel.close();

def count_rows(object):
    rng=object[0].range(object[1]).expand('table');
    rowcount=rng.rows.count;
    return rowcount;

def find_the_end(sheet,key,begin:int,limit:int):
    '''Get the end of a record in history.
    Input: the searching key, the beginning and the maximum line considered.
    Output: the end'''
    i=begin+1;
    while sheet.range('{0}{1}'.format(key,i)).value==None and i<=limit:
        i=i+1;
    return i-1;