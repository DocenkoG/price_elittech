# -*- coding: UTF-8 -*-
import os
import os.path
import logging
import logging.config
import io
import sys
import configparser
import time


global log
global myname

def config_read():
   global myname
   global mydir
   global vendors
 
   cfgFName = os.path.join( mydir, myname + '.cfg')
   config = configparser.ConfigParser()
   if os.path.exists(cfgFName):   
      config.read( cfgFName)
   else : 
      print('Не найден файл конфигурации: '+ cfgFName)
      exit(3)
 
   # в разделе [VENDORS] находится список интересующих нас поставщиков
   temp_list = config.options('VENDORS')
   print(temp_list)
   vendors=[]
   for vName in temp_list :
      if (1 == config.get('VENDORS', vName)) :
         vendors.append(vName)



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

 
   # Прочитать конфигурацию из файла
   cfg = config_read()
   



if __name__ == '__main__':
   global  myname
   global  mydir
   myname   = os.path.basename(os.path.splitext(sys.argv[0])[0])
   mydir    = os.path.dirname (sys.argv[0])
   main( )
 
'''
elittech_price.init()
if  elittech_price.downloader():
    elittech_price.convert2csv()


#os.system(r'c:\prices\_scripts\remove_tmp_profiles.cmd')
'''