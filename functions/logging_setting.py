#-*-coding:utf-8-*-
import logging
import os

class Log():
    
    def __init__(self, logger=None):
        # debug, info, warn, error, critical
        self.logger = logging.getLogger(logger)
        self.logger.setLevel(level=logging.DEBUG)
        formatter = logging.Formatter(fmt='%(asctime)s - %(filename)s[line:%(lineno)d] - %(levelname)s : %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
        # output to log file
        root_path = os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir))
        handler = logging.FileHandler(f'{root_path}/logs/error.log', encoding='utf-8')
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