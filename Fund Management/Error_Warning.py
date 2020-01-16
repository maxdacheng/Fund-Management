'''
Errors and warnings
'''

class Warning:
    '''
    Different kinds of warnings the system could raise
    '''

    def yes_and_no_warning(*message):
        r="";
        while True:
            count=0;
            for m in message:
                if count==0:
                    print('Warning: '+m);
                else:
                    print(m);
                count+=1;
            print("Y.Yes, N.No");
            r=input();
            if r in ['Y','N']:
                break;
        if r=='Y':
            return True;
        else:
            return False;


    def warning(*message):
        count=0;
        for m in message:
            if count==0:
                print('Warning: '+m);
            else:
                print(m);
            count+=1;



