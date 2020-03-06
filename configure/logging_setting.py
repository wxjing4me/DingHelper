#-*-coding:utf-8-*-
import logging
from logging import handlers
from os import pardir as os_pardir
from os.path import abspath as os_path_abspath, join as os_path_join, dirname as os_path_dirname
from re import compile as re_compile

from configure.default_setting import ROOT_DIR

class Log():
    
    def __init__(self, logger=None):
        # debug, info, warn, error, critical
        self.logger = logging.getLogger(logger)
        self.logger.setLevel(level=logging.DEBUG)
        formatter = logging.Formatter(fmt='%(asctime)s - %(filename)s[line:%(lineno)d] - %(levelname)s : %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
        # output to log file
        handler = handlers.TimedRotatingFileHandler(filename=f"{ROOT_DIR}/logs/error", encoding='utf-8', when="midnight", interval=1, backupCount=3)
        handler.suffix = "%Y-%m-%d.log"
        handler.extMatch = re_compile(r"^\d{4}-\d{2}-\d{2}.log$")
        # handler = logging.FileHandler(f'{root_path}/logs/error.log', encoding='utf-8')
        handler.setFormatter(formatter)
        handler.setLevel(level=logging.ERROR)
        self.logger.addHandler(handler)
        # output to console
        console = logging.StreamHandler()
        console.setFormatter(formatter)
        self.logger.addHandler(console)

    def getLog(self):
        return self.logger

if __name__ == '__main__':
    log = Log(__name__).getLog()
    log.info('Main info')
    log.debug('Main Debug')
    log.warn('Main warn')
    log.error('Main Error')
    log.critical('Main critical')