class asset: 
    def __init__(self, type):
        self.type = type;
        self.ID="000000";
        self.name="none";
        self.inivalue=0.0;
        self.curshare=0.0;
        self.ratio=0.0;
        self.inineat=0.0;
        self.curvalue=0.0;
        self.curneat=0.0;
        self.curgrowth=0.0;
        self.lastneat=0.0;
        self.lastgrowth=0.0;
        self.arevenue=0.0;
        self.agrowth=0.0;
        self.row=0;

    def isinvalid(self):
        if self.curneat==0.0 and self.curgrowth==0.0:
            return True;
        else:
            return False;
    
    def isempty(self):
        if self.name=='none':
            return True;
        else:
            return False;

class portfolio:
    def __init__(self):
        self.date="0000/00/00";
        self.time="00:00";
        self.content=list();
        self.pvalue=0.0;
        self.savevalue=0.0;
        self.stockvalue=0.0;
        self.bondvalue=0.0;
        self.monevalue=0.0;
        self.sparevalue=0.0;
        self.prevenue=0.0;
        self.pgrowth=0.0;
        self.inivalue=0.0;

    def release(self):
        #Clear the portfolio.
        #Input: the portfolio object.
        #Output: null.
        self.content.clear();
        self.date="0000/00/00";
        self.time="00:00";
        self.pvalue=0.0;
        self.prevenue=0.0;
        self.pgrowth=0.0;
        self.bondvalue=0.0;
        self.savevalue=0.0;
        self.stockvalue=0.0;
        self.monevalue=0.0;
        self.sparevalue=0.0;
        self.prevenue=0.0;
        self.pgrowth=0.0;
        self.inivalue=0.0;
        

    def isempty(self):
        if self.prevenue==0.0 and len(self.content)==0:
            return True;
        else:
            return False;