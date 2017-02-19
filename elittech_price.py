# -*- coding: utf-8 -*-
'''
Программа для манипуляции с файлами при скачивании прайса поставщика / конкурента.
Файлы находятся в папке, задаваемой параметром.
Если в этой папке присутствует файл с именем 1.txt, то скачивание прайса не делаем.
Если скачиваем архив, то дату файла проверяем после разархивации.
Сверяем дату с предыдущей копией файла. Если не новый файл, то работа завершена.
Новый прайс копируем в "old" для сравнения дат в следующий раз, и в файл для обработки макросом.
После этого делаем файл "1.txt" (признак того, что файл обновлен и требуется его обработка макросом)
'''

import os
import os.path
import shutil
from datetime import datetime
import logging
import logging.handlers


class Vendor(object):
    """Объект поставщика, прайс которого мы скачиваем с сайта. """
    def __init__(self, 
                orgName,                                                     # Имя поставщика (подпапка прайслистов)
                PathPrice = r'\\asgard\приложения\Монополия\Прайслисты',     # Путь к папкам с прайсами поставщиков
                PathUnit  = r'c:\prices' ,                                   # Путь к папкам c юниттестами поставщиков
                PathDwnld = r'c:\prices\downloads'                           # Путь к папке загрузок
                ):
        #super(Vendor, self).__init__()
        self.orgName   = orgName
        self.PathPrice = PathPrice
        self.PathDwnld = PathDwnld
        self.PathUnit  = PathUnit
        logMain = 'get_price.log'
        logOrg  = 'get_price_'+orgName+'.log'
        self.FlogMain = os.path.join( PathPrice, logMain)
        self.FlogOrg  = os.path.join( PathPrice, self.orgName, logOrg )
        self.make_loger()



    def make_loger(self):
        # Создать регистратор верхнего уровня
        self.loger= logging.getLogger(self.orgName)
        subject = 'Subj_for_alert_mail, ' + self.orgName 
        self.loger.setLevel(logging.DEBUG)

        # Добавить общий обработчик и частный для поставщика
        handlrMain =logging.handlers.RotatingFileHandler(self.FlogMain, 'a', 512000, 5)
        handlrOrg  =logging.handlers.RotatingFileHandler(self.FlogOrg,  'a', 128000, 3)
        handlrMail =logging.handlers.SMTPHandler(('192.168.10.3',25), 'docn@meijin.ru', 'kissupport@meijin.ru', subject,  credentials=None, secure=None)
        handlrMail.setLevel(logging.CRITICAL) 
        self.loger.addHandler(handlrMain)
        self.loger.addHandler(handlrOrg)
        self.loger.addHandler(handlrMail)

        # Описать форматы для обработчиков
        fmt1    = logging.Formatter('%(asctime)s  %(levelname)-7s %(message)s')
        fmtMain = logging.Formatter('%(asctime)s  %(levelname)-7s'+ self.orgName +' %(message)s')
        fmtOrg  = logging.Formatter('%(asctime)s  %(levelname)-7s %(message)s')
        handlrMain.setFormatter(fmtMain)
        handlrOrg.setFormatter( fmtOrg)
        handlrMail.setFormatter(fmtOrg)


    def convert2csv(self):
        # вызвать скрипт преобразования csv, если такой существует
        self.loger.debug('Вызван метод "convert2csv()"')
        #
        script2csvName = os.path.join( self.PathPrice, self.orgName, self.orgName +'_auto.py')
        if  not os.path.exists(script2csvName):
            self.loger.info( r'Отсутствует скрипт для преобразования в CSV: ' + script2csvName)
        else:
            self.loger.info( r'Вызываем скрипт для преобразования в CSV: ' + script2csvName)
            work_dir = os.getcwd()                                                  
            os.chdir( os.path.join( self.PathPrice, self.orgName ))
            os.system(r'python.exe ' + script2csvName + ' ' + self.orgName)             # Вызов 
            os.chdir(work_dir)



    def get_file(self):
        self.loger.info('.')
        f1txt   = os.path.join( self.PathPrice, self.orgName, r'1.txt')   
        result  = False
        FunitName = os.path.join( self.PathUnit, self.orgName, self.orgName +'_unittest.py')
        if  not os.path.exists(FunitName):
            self.loger.debug( r'Отсутствует юниттест для загрузки прайса ' + FunitName)
        else:
            dir_befo_download = set(os.listdir(self.PathDwnld))
            os.system(r'c:\Python27\python.exe ' + FunitName)                               # Вызов unittest'a
            dir_afte_download = set(os.listdir(self.PathDwnld))
            new_files = list( dir_afte_download.difference(dir_befo_download))
            if len(new_files) == 1 :   
                new_file = new_files[0]                                                     # загружен ровно один файл. 
                new_ext  = os.path.splitext(new_file)[-1]
                new_name = os.path.splitext(new_file)[0]
                DnewFile = os.path.join( self.PathDwnld,new_file)
                new_file_date = os.path.getmtime(DnewFile)
                self.loger.debug( r'Скачанный файл ' +DnewFile + r' имеет дату ' + str(datetime.fromtimestamp(new_file_date) ) )
                if new_ext == '.zip':                                                       # Архив. Обработка не завершена
                    self.loger.debug( r'Zip-архив. Разархивируем.')
                    work_dir = os.getcwd()                                                  
                    os.chdir( os.path.join( self.PathUnit, self.orgName ))
                    dir_befo_download = set(os.listdir(os.getcwd()))
                    os.system(r'unzip -o ' + DnewFile)
                    os.remove(DnewFile)   
                    dir_afte_download = set(os.listdir(os.getcwd()))
                    new_files = list( dir_afte_download.difference(dir_befo_download))
                    if len(new_files) == 1 :   
                        new_file = new_files[0]                                             # разархивирован ровно один файл. 
                        new_ext  = os.path.splitext(new_file)[-1]
                        DnewFile = os.path.join( os.getcwd(),new_file)
                        new_file_date = os.path.getmtime(DnewFile)
                        self.loger.debug( r'Файл из архива ' +DnewFile + r' имеет дату ' + str(datetime.fromtimestamp(new_file_date) ) )
                        DnewPrice = DnewFile
                    elif len(new_files) >1 :
                        self.loger.debug( r'В архиве не единственный файл. Надо разбираться.')
                        DnewPrice = "dummy"
                    else:
                        self.loger.debug( r'Нет новых файлов после разархивации. Загляни в папку юниттеста поставщика.')
                        DnewPrice = "dummy"
                    os.chdir(work_dir)

                elif new_ext == '.csv' or new_ext == '.htm' or new_ext == '.xls' or new_ext == '.xlsx':
                    DnewPrice = DnewFile                                                    # Имя скачанного прайса

                if DnewPrice != "dummy" :
                    FoldName = os.path.join(self.PathPrice,self.orgName, 'old_'+self.orgName+new_ext) # Предыдущая копия прайса, для сравнения даты
                    FnewName = os.path.join(self.PathPrice,self.orgName, 'new_'+self.orgName+new_ext) # Файл, с которым работает макрос

                    if  (not os.path.exists( FnewName)) or new_file_date>os.path.getmtime(FnewName) : 
                        self.loger.debug( r'Предыдущего прайса нет или он устарел. Копируем новый.' )
                        if os.path.exists( FoldName): os.remove( FoldName)
                        if os.path.exists( FnewName): os.rename( FnewName, FoldName)
                        shutil.copy2(DnewPrice, FnewName)
                        os.system(r'echo Hello > ' + f1txt)
                        result = True
                    else:
                        self.loger.debug( r'Предыдущий прайс не старый, копироавать не надо.' )

                    # Убрать скачанные файлы
                    if  os.path.exists(DnewPrice):  os.remove(DnewPrice)   
                    
            elif len(new_files) == 0 :        
                self.loger.debug( r'Не удалось скачать файл прайса ')
            else:
                self.loger.debug( r'Скачалось несколько файлов. Надо разбираться ...')
        self.loger.info( r'End' )
        #self.loger.removeHandler(hOrg)
        #self.loger.removeHandler(hMain)
        return result
