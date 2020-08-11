import datetime
from os.path import exists
from os import makedirs
import traceback

def logErrors(func):

    def wrapper(*args, **kwargs):

        try:
            return func(*args, **kwargs)
        except Exception as exc:
            now = datetime.datetime.now()
            if not exists(r'.\BOlogfiles'):
                makedirs(r'.\BOlogfiles')
            logFile = open(r'.\BOlogfiles\BO_logfile_%s.txt' %'_'.join([str(now.year), str(now.month), str(now.day), str(now.hour), str(now.minute), str(now.second)]),
                           'w')
            logFile.write(traceback.format_exc())
            logFile.close()
            print('Exception written to file')
            raise Exception(exc)

    return wrapper
