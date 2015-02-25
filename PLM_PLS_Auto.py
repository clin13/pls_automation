#!/usr/bin/env python
# encoding: utf-8

import sys
import os
import urllib2
import zipfile
import ConfigParser
import shutil
import time
import datetime
import re
import csv
import xlwt
import logging

__version__ = 0.1
__date__ = '2015-01-19'
INICount = 0
CmpntLstCount = 0

def MergeCSV(csv_folder):
    book = xlwt.Workbook()
    for fil in os.listdir(csv_folder):
        sheet = book.add_sheet(fil[:-4])    # File name only
        with open(csv_folder + "/" + fil) as filename:
            res = csv.reader(filename)
            i = 0
            for row in res:
                for j, each in enumerate(row):
                    sheet.write(i, j, each) # i - row, j - column, each - value
                i += 1
    book.save(csv_folder[:-11]+"PLM_PLS_List.xls")

def GetPLSVersion(product, INI_FILE, pls_path, i):
    Language_ID = ("1", "32", "128", "64", "512", "8192", "4096", "16", "2", "4", "8")
    PLM_PLS_ID = ("141", "696", "177", "189", "462", "490", "539", "523", "602", "594")
    Component_name = ("PLM 2.0", "PLM 2.1", "TMMS", "IDF", "TMSM", "VDI", "Toolbox", "DLP", "TMEE", "TMEAC")
    Language_name = ("EN", "DE", "ES", "FR", "IT", "RU", "PO", "KO", "TC", "JP", "CN")
    L = 0                                   # Language index

    Cfg = ConfigParser.ConfigParser()
    Cfg.read(INI_FILE)

    for L_ID in Language_ID:
        # Create PLM/PLS list files with version by language
        f_name = pls_path+'/'+Language_name[L]+'.csv'
        if (os.path.exists(f_name)):
            fd = open(f_name, 'ab+')
        else:
            fd = open(f_name, 'wb+')

        L += 1
        f = csv.writer(fd, delimiter = ",", lineterminator="\n", quoting=csv.QUOTE_MINIMAL)

        # i = 0 - EN, EMEA, TC, KO, i = 1 - JP, i = 2 - CN
        if (i == 0 and L_ID != "4" and L_ID != "8") or (i == 1 and L_ID == "4") or (i == 2 and L_ID == "8"):
            f.writerow([product])
            f.writerow(["Component name", "AU ID", "Version"])
            j = 0

            for pls_id in PLM_PLS_ID:
                Section2 = ''
                if pls_id == '177':         # TMMS
                    Section = 'Info_%s_50000_%s_1'%(pls_id, L_ID)
                    if (Cfg.has_section(Section)):
                        Section = 'Info_%s_50000_%s_1'%(pls_id, L_ID)
                        # OSCE 10 FR TMMS has two records
                        if (L_ID == "64" and Cfg.has_section("Info_177_50000_64_1")
                            and Cfg.has_section("Info_177_55000_64_1")):
                            Section2 = 'Info_177_55000_64_1'
                    else:
                        Section = 'Info_%s_55000_%s_1'%(pls_id, L_ID)
                else:
                    Section = 'Info_%s_10000_%s_1'%(pls_id, L_ID)
                    if (pls_id == '189' and not Cfg.has_section(Section)):
                        Section = 'Info_%s_10000_%s_1'%("190", L_ID)
                    if (pls_id == '539' and not Cfg.has_section(Section)):
                        Section = 'Info_%s_10000_%s_1'%("540", L_ID)
                    if (pls_id == '602' and not Cfg.has_section(Section)):
                        Section = 'Info_%s_10000_%s_1'%("603", L_ID)
                    if (pls_id == '594' and not Cfg.has_section(Section)):
                        Section = 'Info_%s_10000_%s_1'%("595", L_ID)

                if (Cfg.has_section(Section)):
                    ver = Cfg.get(Section, "Version")
                    build = Cfg.get(Section, "Build")
                    # OSCE 10 FR TMMS only
                    if (Cfg.has_section(Section2)):
                        ver2 = Cfg.get(Section2, "Version")
                        build2 = Cfg.get(Section2, "Build")
                        version = ver+'.'+build+' / '+ver2+'.'+build2
                    else:
                        version = ver+'.'+build
                else:
                    version = 'N/A'

                if product == 'osce10' and j == 0:
                    f.writerow(["PLM 1.0", pls_id, version])
                else:
                    f.writerow([Component_name[j], pls_id, version])
                j+=1
            f.writerow([])
        fd.close()
        time.sleep(0.3)

def GetComponentList(prdct, INI, dir, i, logger):
    cfg = ConfigParser.ConfigParser()
    cfg.read(INI)
    loop = 1
    loop_enable = 0
    global CmpntLstCount
    # Both PLM 2.0 & 2.1 exist
    if (cfg.has_option("All_Product", "Product.697")):
        if (cfg.has_option("All_Product", "Product.138")):
            loop = 2
            loop_enable = 1
            CL_AU_ID = '0'
        else:
            CL_AU_ID = '697'                # PLM 2.1 AU ID
    else:
        CL_AU_ID = '138'

    if i == 0:
        language_ID = ("1", "32", "128", "64", "512", "8192", "4096", "16", "2")
    elif i == 1:
        language_ID = '4'
    else:
        language_ID = '8'

    AU_svr = cfg.get("Server", "Server.1")  # Get AU server URL

    while(loop > 0):
        if (loop_enable):
            if loop==2:
                CL_AU_ID = '697'
            else:
                CL_AU_ID = '138'
        for lang_ID in language_ID:
            section = 'Info_%s_10000_%s_1'%(CL_AU_ID, lang_ID)

            if (cfg.has_section(section)):
                path, size = cfg.get(section, "PATH").split(',')
                full_path = AU_svr+'/'+path
                # '.' - Any character except a new line, '+' - 1 or more repetitions, '?' - 0 or 1 repetition
                pattern = prdct+'/(.+?)/'
                language = re.search(pattern, path)
                # Group(1) - string searched by regular expression
                if (CL_AU_ID == "697"):
                    os.makedirs(os.path.join(dir, "21_"+language.group(1)))
                    lang_dir = "21_"+language.group(1)
                else:
                    os.makedirs(os.path.join(dir, language.group(1)))
                    lang_dir = language.group(1)

                # Ex: Length of product/osce11/cht/
                zipfile = path[10+len(prdct)+len(language.group(1)):]
                zipfile_path = dir+'/'+lang_dir+'/'+zipfile
                buf = ''

                try:
                    buf = urllib2.urlopen(full_path)  # Download component list
                except urllib2.URLError as err:
                    msg = 'Can not download '+zipfile_path+'\n'
                    print str(err.reason)
                    print msg
                    logger.debug(str(err.reason))
                    logger.debug(msg)
                    return
                except urllib2.HTTPError as err:
                    msg = 'Can not download '+zipfile_path+'\n'
                    print str(err.code)
                    print msg
                    logger.debug(str(err.reason))
                    logger.debug(msg)
                    return

                relative_path = prdct+'/'+lang_dir+'/'+zipfile
                print 'Download '+relative_path
                try:
                    file = open(zipfile_path, 'wb')
                except IOERROR:
                    logger.debug("Can not open "+relative_path)
                time.sleep(0.4)
                data = buf.read()
                time.sleep(0.5)
                if not data:
                    logger.debug("Write "+relative_path+" fail")
                    return
                if (data):
                    file.write(data)
                else:
                    data = buf.read()
                    file.write(data)
                    print '..Retry writing data'
                time.sleep(1)               # In case of file not written successfully
                CmpntLstCount+=1
                file.close()
                buf.close()
                data = ''
                time.sleep(0.5)
        loop-=1

def GetServerINI(p, p_URL, dir, i, pls_path, logger):
    if i == 0:
        lan = '_'                           # EN, EMEA, TC, KO
    elif i == 1:
        lan = '_jp_'
    else:
        lan = '_cn_'

    try:
        res = urllib2.urlopen(p_URL)        # Open URL of AU's server.ini
    except urllib2.URLError as e:
        Reason = str(e.reason)
        msg = 'Can not download '+p+lan+'server.ini\n'
        print Reason
        print msg
        logger.debug(Reason)
        logger.debug(msg)
        return
    except urllib2.HTTPError as e:
        msg = 'Can not download '+p+lan+'server.ini\n'
        print str(e.code)
        print msg
        logger.debug(str(e.code))
        logger.debug(msg)
        return

    INI_FILE = dir+'/'+p+lan+'server.ini'   # Rename to product_language_server.ini
    f = open(INI_FILE, 'wb')
    data = res.read()
    time.sleep(0.5)
    if not data:
        logger.debug("Write "+INI_FILE+" fail")
        return
    f.write(data)
    time.sleep(0.8)
    f.close()
    res.close()
    global INICount
    INICount+=1

    print "Download %s%s"%(p, lan)+"server.ini"
    GetComponentList(p, INI_FILE, dir, i, logger)

def Init(argv=None):
    print '--------------------------------------------------------------'
    print 'PLM_PLS_Automation v'+str(__version__)+'  Release date: '+__date__
    print '--------------------------------------------------------------'
    now = datetime.datetime.now()           # Get current time
    str_now = str(now.year)+"_"+str(now.month)+"_"+str(now.day)+"_"+str(now.hour)+"_"+str(now.minute)+"_"+str(now.second)
    root_dir = os.path.join(os.path.abspath(os.getcwd()), str_now+"_"+"Component_Lists")
    if os.path.exists(root_dir):            # If folder exists, delete it
        shutil.rmtree(root_dir)
        time.sleep(1.0)
    os.makedirs(root_dir)
    os.makedirs(os.path.join(root_dir, "pls_by_lang"))
    pls_path = os.path.join(root_dir, "pls_by_lang")
    log_name = root_dir+'/'+'pls.log'       # Debug log
    logging.basicConfig(filename=log_name, level=logging.DEBUG)
    logger = logging.getLogger(__name__)

    product = ("osce10", "osce105", "osce106", "osce106sp2", "osce11", "cm55", "cm60")
    URL_head = ("http://")
    URL_tail = (
                "-p.activeupdate.trendmicro.com/activeupdate/server.ini",
                "-p.activeupdate.trendmicro.co.jp/activeupdate/japan/server.ini",
                "-p.activeupdate.trendmicro.com.cn/activeupdate/china/server.ini")
    lang_set = ("_server.ini", "_jp_server.ini", "_cn_server.ini")

    for p in product:
        tmp_root = root_dir
        os.makedirs(os.path.join(tmp_root, p))
        i = 0
        for t in URL_tail:
            p_URL = URL_head + p + t
            GetServerINI(p, p_URL, os.path.join(tmp_root, p), i, pls_path, logger)
            i+=1
    print '\nCreating PLS version files by language'

    for prdct in product:
        tmp = root_dir
        prdct_path = os.path.join(tmp, prdct)
        j = 0
        for lang in lang_set:
            INI_path = prdct_path+'/'+prdct+lang
            GetPLSVersion(prdct, INI_path, pls_path, j)
            j += 1
        
    print '\nMerge all csv files to a Excel file'
    MergeCSV(pls_path)

    global CmpntLstCount
    global INICount
    print '\nServer.ini:      '+str(INICount)+' file(s)'
    print 'Component list:  '+str(CmpntLstCount)+' file(s)\n'

if __name__ == '__main__':
    sys.exit(Init())
