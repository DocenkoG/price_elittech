# -*- coding: UTF-8 -*-
import os
import os.path
import logging
import logging.config
import io
import sys
import configparser
import time
import elittech_downloader
import elittech_converter

global log
global myname



def make_loger():
    global log
    logging.config.fileConfig('logging.cfg')
    log = logging.getLogger('logFile')
    log.debug('test debug message')
    log.info( 'test info message')
    log.warn( 'test warn message')
    log.error('test error message')
 #  log.critical('test critical message')



def main( ):
    global  myname
    global  mydir
    print('myname    =', myname)
    print('mydir     =', mydir)
   
    make_loger()
    log.debug(myname +', Begin main.')

    if  elittech_downloader.download() :
        elittech_converter.convert2csv()



if __name__ == '__main__':
    global  myname
    global  mydir
    myname   = os.path.basename(os.path.splitext(sys.argv[0])[0])
    mydir    = os.path.dirname (sys.argv[0])
    main( )
 
#os.system(r'c:\prices\_scripts\remove_tmp_profiles.cmd')