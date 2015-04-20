#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sqlite3
import time
import sys
import getopt
import requests
from lxml import html
import xlsxwriter

_version    = '1.0.2'
_file_name  = "zakupki_gov_parser"

_verbose    = False
_log_write  = False
_action     = ""

_sleep_time  = 1

_links_path     = 'links.txt'
_db_path        = 'db.sql'
_log_path       = 'zakupki_log.txt'
_export_path    = 'export.xlsx'

_web_headers = { 'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64; rv:37.0) Gecko/20100101 Firefox/37.0' }

_export_headers = ( u"Полное нaименование", u"Сокращенное наименование", u"Телефон", u"Факс", u"Почтовый адрес", u"Email", u"Контактное лицо" )

# Логирование сообщения
def printLog( msg ):
    print "[%s]  %s" % ( time.strftime("%Y-%m-%d %H:%M:%S"), msg )
    if _log_write:
        f = open( _log_path, 'a' )
        f.write( "[%s]  %s\n" % ( time.strftime("%Y-%m-%d %H:%M:%S"), msg ) )
        f.close()
    
def printLogVerbose( msg ):
    if _verbose:
        print "[%s]  %s" % ( time.strftime("%Y-%m-%d %H:%M:%S"), msg )
        if _log_write:
            f = open( _log_path, 'a' )
            f.write( "[%s]  %s\n" % ( time.strftime("%Y-%m-%d %H:%M:%S"), msg ) )
            f.close()

# Вывод списка параметров действий
def printActionList():
    print "Action list:"
    print "  [*] links  - parse links from page to db;"
    print "  [*] parse  - parse data to db;"
    print "  [*] export - export data from db to xlsx file"

# Создание таблиц в файле базы
def createTables( conn ):
    conn.execute( '''CREATE TABLE IF NOT EXISTS `links` (
                `id`                   INTEGER PRIMARY KEY,
                `url`                  VARCHAR(255) UNIQUE,
                `res`                  INTEGER DEFAULT 0
                )''' )
    
    conn.execute( '''CREATE TABLE IF NOT EXISTS `data` (
                `id`                   INTEGER PRIMARY KEY,
                `full_name`            VARCHAR(255) NOT NULL,
                `short_name`           VARCHAR(255),
                `phone`                VARCHAR(100),
                `fax`                  VARCHAR(255),
                `address`              VARCHAR(255),
                `email`                VARCHAR(255),
                `name`                 VARCHAR(255)
            )''')

# Запрос страницы
def getPage( url ):
    global _web_headers
    response = requests.get( url, headers = _web_headers )
    if response.status_code != 200:
        return False
    else:
        return response

# Парсинг ссылок со страницы 
def parseLinks( url, sqlConn ):
    global _verbose
    response = getPage( url )
    parsed_body = html.fromstring( response.text )
    links_count = 0
    
    for url in parsed_body.xpath( "//dl//dt//a/@href" ):
        try:
            sqlConn.execute( "INSERT INTO `links`(url) VALUES('%s')" % ( url ) )
            links_count = links_count + 1
            
            printLogVerbose( "Add link %s to DataBase" % ( 'http://zakupki.gov.ru' + url ) )
            
        except sqlite3.IntegrityError as e:
            pass
            

    sqlConn.commit()
    printLog( "Added %d links from page" % ( links_count ) )
    
def getShortName( url ):
    response = getPage( url )
    if response != False:
        parsed_body = html.fromstring( response.text )
    
        if len( parsed_body.xpath( "//td//span[ contains( text(), '" + u'Сокращенное наименование' + "' ) ]//..//..//td[position() = 2]//span/text()" ) ) > 0:
            return parsed_body.xpath( "//td//span[ contains( text(), '" + u'Сокращенное наименование' + "' ) ]//..//..//td[position() = 2]//span/text()" )[0].strip()
        else:
            return ""
    else:
        printLog( "Url %s load filed" % ( url ) )
        return False
    
def parseOrg( urlId, url, sqlConn ):
    global _web_headers
    
    response = getPage( 'http://zakupki.gov.ru' + url )
    if response != False:
        parsed_body = html.fromstring( response.text )
        
        _orgName        = ""
        _orgShortName   = ""
        _orgPhone       = ""
        _orgFax         = ""
        _orgAddress     = ""
        _orgEmail       = ""
        _orgFio         = ""
        
        # Название организации
        if len( parsed_body.xpath( "//td[ contains( text(), '" + u'Наименование организации' + "' ) ]//..//td[position() = 2]/text()" ) ) > 0:
            _orgName = parsed_body.xpath( "//td[ contains( text(), '" + u'Наименование организации' + "' ) ]//..//td[position() = 2]/text()" )[0].strip()
        elif len( parsed_body.xpath( "//td[ contains( text(), '" + u'Организация, осуществляющая закупку' + "' ) ]//..//td[position() = 2]/text()" ) ) > 0:
            _orgName = parsed_body.xpath( "//td[ contains( text(), '" + u'Организация, осуществляющая закупку' + "' ) ]//..//td[position() = 2]/text()" )[0].strip()
        elif len( parsed_body.xpath( "//td[ contains( text(), '" + u'Организация осуществляющая закупку' + "' ) ]//..//td[position() = 2]/text()" ) ) > 0:
            _orgName = parsed_body.xpath( "//td[ contains( text(), '" + u'Организация осуществляющая закупку' + "' ) ]//..//td[position() = 2]/text()" )[0].strip()
        else:
            return False
        
        # Номер телефона
        if len( parsed_body.xpath( "//td[ contains( text(), '" + u'Номер контактного телефона' + "' ) ]//..//td[position() = 2]/text()" ) ) > 0:
            _orgPhone = parsed_body.xpath( "//td[ contains( text(), '" + u'Номер контактного телефона' + "' ) ]//..//td[position() = 2]/text()" )[0].strip()
            
        # Номер факс
        if len( parsed_body.xpath( "//td[ contains( text(), '" + u'Факс' + "' ) ]//..//td[position() = 2]/text()" ) ) > 0:
            _orgFax = parsed_body.xpath( "//td[ contains( text(), '" + u'Факс' + "' ) ]//..//td[position() = 2]/text()" )[0].strip()
            
        # Почтовый адрес
        if len( parsed_body.xpath( "//td[ contains( text(), '" + u'Почтовый адрес' + "' ) ]//..//td[position() = 2]/text()" ) ) > 0:
            _orgAddress = parsed_body.xpath( "//td[ contains( text(), '" + u'Почтовый адрес' + "' ) ]//..//td[position() = 2]/text()" )[0].strip()
            
        # Контактный Email
        if len( parsed_body.xpath( "//td[ contains( text(), '" + u'Адрес электронной почты' + "' ) ]//..//td[position() = 2]/text()" ) ) > 0:
            _orgEmail = parsed_body.xpath( "//td[ contains( text(), '" + u'Адрес электронной почты' + "' ) ]//..//td[position() = 2]/text()" )[0].strip()
        
        # ФИО ответственного лица
        if len( parsed_body.xpath( "//td[ contains( text(), '" + u'Ответственное должностное лицо' + "' ) ]//..//td[position() = 2]/text()" ) ) > 0:
            _orgFio = parsed_body.xpath( "//td[ contains( text(), '" + u'Ответственное должностное лицо' + "' ) ]//..//td[position() = 2]/text()" )[0].strip()
        
        # Короткое название
        if len( parsed_body.xpath( "//td//a/@href" ) ) > 0:
            shortUrl        = parsed_body.xpath( "//td//a/@href" )[0].strip()
            _orgShortName   = getShortName( shortUrl )
            if _orgShortName == False:
                return False
            
        sqlConn.execute( "INSERT INTO `data` VALUES(%d, '%s', '%s', '%s', '%s', '%s', '%s', '%s')" % ( urlId, _orgName, _orgShortName, _orgPhone, _orgFax, _orgAddress, _orgEmail, _orgFio ) )
        sqlConn.commit()
        
        return True
    
    else:
        printLog( "Url %s load filed" % ( url ) )
        return False

def main():
    global _verbose
    global _db_path
    global _links_path
    global _action
    global _log_write
    global _log_path
    global _export_path
    global _export_headers
    global _sleep_time
    
    # Разбор параметров командной строки
    options, remainder = getopt.getopt( sys.argv[1:], 'vd:l:s:', [
                                                                'verbose',
                                                                'version',
                                                                'dbfile=',
                                                                'linksfile=',
                                                                'action=',
                                                                'logfile=',
                                                                'exportfile=',
                                                                'sleep='
                                                            ] )

    for opt, arg in options:
        if opt in ( '-v', '--verbose' ):
            _verbose = True

        elif opt == '--version':
            print "zakupki.gov parser version: %s" % ( _version )
            exit(0)

        elif opt in ( '-d', '--dbfile' ):
            _db_path = arg

        elif opt == '--action':
            if arg == "list":
                printActionList()
                exit(0)
            elif arg == "links":
                _action = "links"
            elif arg == "parse":
                _action = "parse"
            elif arg == "export":
                _action = "export"
            else:
                print "Unkonwn action. Using: %s --action list" % ( _file_name )

        elif opt in ( '-l', '--linksfile' ):
            _links_path = arg
        
        elif opt == '--logfile':
            _log_write = True
            _log_path = arg
         
        elif opt == '--exportfile':
            _export_path = arg
            
        elif opt in ( '-s', '--sleep' ):
            _sleep_time = int( arg )

    # Соединение с базой данных
    sqlConn = sqlite3.connect( _db_path )
    printLogVerbose( "Connect to db file %s" % ( _db_path ) )
        
    # Опредение действия
    if _action == "links":
        createTables( sqlConn )
         
        f = open( _links_path )
        for url in f:
            printLogVerbose( "parse link from %s" % ( url[:-1] ) )
                 
            parseLinks( url, sqlConn )
             
            printLogVerbose( "from %s links parsed" % ( url[:-1] ) )
             
            time.sleep( _sleep_time )
             
        _action = 'parse'
        
    if _action == 'parse':
        cursor = sqlConn.execute( "SELECT `id`, `url` FROM `links` WHERE `res` = '0'" )
        rows = cursor.fetchall()
        for row in rows:
            if parseOrg( row[0], row[1], sqlConn ):
                sqlConn.execute( "UPDATE `links` SET `res` = 1 WHERE `id` = %d" % ( row[0] ) )
                sqlConn.commit()
                printLog( "Parse %s success . . ." % ( 'http://zakupki.gov.ru' + row[1] ) )
            else:
                printLog( "Parse %s failed . . ." % ( 'http://zakupki.gov.ru' + row[1] ) )
                
            time.sleep( _sleep_time )
            
        _action = 'export'
            
    if _action == 'export':
        cursor = sqlConn.execute( "SELECT `full_name`, `short_name`, `phone`, `fax`, `address`, `email`, `name` FROM data" )
        rows = cursor.fetchall()
        
        workbook = xlsxwriter.Workbook( _export_path )
        worksheet = workbook.add_worksheet()
        
        # Заголовки
        for ( i, header ) in enumerate( _export_headers ):
            worksheet.write( 0, i, header )
            
        for ( i, row ) in enumerate( rows ):
            for j in xrange( len( rows ) ):
                worksheet.write( i + 1, j, row[j] )
                
        workbook.close()
    
    sqlConn.close()

if __name__ == '__main__':
    main()