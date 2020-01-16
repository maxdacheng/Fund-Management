import Global;
import Excel_Processing as Excel;
from Basic_Class import portfolio;
from Main import update,updateintro,record,prints,calculate,plot,help,auto_backup;
from Error_Warning import Warning;


if __name__ ==  '__main__':  
    #Load excel.
    Global.init();
    Excel.init();

    #Establish cache.
    P=portfolio();
    L=list();
    page=0;
    old_command='';

    #Automatic backup check
    if auto_backup():
        pass;
    else:
        print("Automatic backup failing!");

    while(1):
        #Orders.
        if old_command=='':
            command=input(">>>");
        else:
            command=old_command;
        
        #Create a variable to record the result of the current command, where 0 means failure, 1 means success and -1 means an unknown command.
        result=-1;

        #Refresh the number of rows.   
        Excel.count();

        if command == "update":
            result=update(page,P);

        elif command == "record":
            result=record(P);
        
        elif command == "updateintro":
            result=updateintro(P);

        elif command == "exit":
            Excel.close_excel();
            exit();

        elif command == "clear":
            P.release();

        elif command[:5]=='print':
            result=prints(command,P);

        elif command[:4]=='plot':
            result=plot(command);

        elif command[:4]=='help':
            result=help(command);

        else:
            result=calculate(command);

        #Judge whether the operation is a success, failure or an unknown
        if result==1:
            old_command='';
        elif result==0:
            if Warning.yes_and_no_warning("For some reason, your command is not correctly implemented. Would you like to do it again?"):
                old_command=command;
            else:
                old_command='';
        else:
            print('Invalid syntax: Unknown command.');
            old_command='';



