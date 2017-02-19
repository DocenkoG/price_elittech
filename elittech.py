# -*- coding: UTF-8 -*-
import os
import os.path
import logging
import logging.config
import io
import sys
import configparser
import time



def config_read():
  global loger
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
  global loger
  global myname
  global mydir
  # Создать регистратор верхнего уровня
  loger= logging.getLogger(myname)
  logFileName = os.path.join( mydir, myname +'.log')
  subject = 'Subj_for_alert_mail, ' + myname 
  loger.setLevel(logging.DEBUG)
  handlr1  =logging.handlers.RotatingFileHandler(logFileName, 'a', 128000, 5)
  handlr2  =logging.handlers.SMTPHandler(('192.168.10.3',25), 'docn@meijin.ru', 'kissupport@meijin.ru', subject, credentials=None, secure=None)
  handlr2.setLevel(logging.CRITICAL) 
  loger.addHandler(handlr1)
  loger.addHandler(handlr2)
  #loger.propagate = True
  # Описать форматы для обработчиков
  fmt1 = logging.Formatter('%(asctime)s  %(levelname)-7s %(message)s')
  handlr1.setFormatter(fmt1)
  handlr2.setFormatter(fmt1)



def main( ):
  global  myname
  global  mydir
  print('myname    =', myname)
  print('mydir     =', mydir)
  
  logging.config.fileConfig('logging.cfg')
  log = logging.getLogger('logFile')
  log.debug('debug message ..................................................................')
  log.info('info message ....................................................................')
  log.warn('warn message ....................................................................')
  log.error('error message ....................................................................')
#  log.critical('critical message')

  # Прочитать конфигурацию из файла
  cfg = config_read()
  
  # Создать регистратор (loger)
#  make_loger()
#  loger.info(myname +', Begin main')



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