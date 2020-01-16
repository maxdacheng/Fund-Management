'''
Global variables.
'''

import configparser;

config=configparser.ConfigParser();
config.read("Setting.ini");


def init():
    global _global_dict
    _global_dict = {}

    #The command set.
    rule={'print':
        {'date':'P.date',
        'time':'P.time',
        'value':'P.pvalue',
        'yield':'P.prevenue',
        'yieldrate':'P.pgrowth',
        'bondvalue':'P.bondvalue',
        'stockvalue':'P.stockvalue',
        'monevalue':'P.monevalue',
        'savevalue':'P.savevalue'},

         'plot':
         {'date':1,
        'value':11,
        'yield':12,
        'yieldrate':13,
        'bondvalue':'P.bondvalue',
        'stockvalue':'P.stockvalue',
        'monevalue':'P.monevalue',
        'savevalue':'P.savevalue'},
         
         'help':
         {'update':'update the real-time holding of your portfolio',
          'record':'record the holding information in the database',
          'updateintro':'update the information in the introduction page',
          'clear':'release the cache',
          'print':'print relevant holding information according to the keyword',
          'plot':'visualize relevant holding or historical data according to the keyword',
          'exit':'save the changes and close the program'
             }};
    
    _global_dict.update({"rule":rule,"config":config});


def set_value(key,value):
    _global_dict[key] = value


def get_value(key,defValue=None):
    try:
        return _global_dict[key]
    except KeyError:
        return defValue