import string;
import Global;
from Basic_Class import portfolio;

def calculator(s):
    words=string.ascii_letters;
    for a in s:
        if a in words:
            return False;
    try:
        print(eval(s));
        return True;
    except:
        return False;

def printout(s:str,P:portfolio):
    content=s[5:];
    content=content.strip("() ");
    rule=Global.get_value('rule');
    if content=='intro':
        for key,value in rule['print'].items():
            print(key+': '+str(eval(value)));
        return True;
    if content not in rule['print'].keys():
        return False;
    print(eval(rule['print'][content]));
    return True;

def print_help(d):
    for key,value in d.items():
        if key in ['print','plot']:
            print('{0} [keyword] -- {1}'.format(key,value));
        else:
            print('{0} -- {1}'.format(key,value));

def help(command):
    rule=Global.get_value('rule');
    command=command[4:];
    command=command.strip('() ');
    if command=='':
        print_help(rule['help']);
        return True;
    elif command in rule['help'].keys():
        print_help({command:rule['help'][command]});
        return True;
    else:
        return False;

