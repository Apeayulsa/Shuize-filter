# -*- coding: utf-8 -*-
import xlrd
from os import popen
from sys import argv

def com(arg,arg1):
    if arg == '-h' or arg == "" or arg1 == '':
        print("\n"+popen('date').read()+"[+] Usage: python Shuize-filter.py [parameter] <target_folder>")
        print("[+] python Shuize-filter.py -k ../../result/\n")
        print('-t\t\tTraverse the current folder')
        print('-h\t\tHelp')
        print('-k\t\tRecursion\n')
        exit()
    if arg == "-t":
        ap = popen('ls %s' % (arg1)).read()
        main(ap,arg1)
    if arg == '-k':
        ap = popen('find %s' % (arg1)).read()
        main(ap,arg1,1)
        

def main(ap,path,flag=0):
    if path[-1] != '/': path+="/"
    popen('mkdir -p %sresult-new/issue' % (path))
    ap = ap.split('\n')[:-1]
    if ap == "": exit("[x] Value_Null")
    for i in ap:
        if i[-4::] != "xlsx": continue
        if flag == 1:
            name = ((i[::-1])[:i[::-1].find('/')])[::-1]
            pathL = i.replace(name,'')
            obj = xlrd.open_workbook(pathL+name)
            tail(obj,name,path,pathL)
        else:
            obj = xlrd.open_workbook(path+i)
            tail(obj,i,path)
        

def tail(obj,name,path,pathL=""):
    if pathL == "": pathL=path
    try:
        sheet = obj.sheet_by_name(u'\u5b58\u6d3b\u7f51\u7ad9\u6807\u9898')
    except:
        return
    status_code = sheet.col_values(1)
    if obj.sheet_names()[-2] != u'\u5b58\u6d3b\u7f51\u7ad9\u6807\u9898':
        popen('mv %s%s %sresult-new/issue' % (pathL,name,path))
        print("%s\t Allowed in issue!" % (name))
        return
    if status_code.count(200) < 2: return
    popen('mv %s%s %sresult-new/' % (pathL,name,path))
    print("%s\t Allowed!" % (name))

if __name__ == "__main__":
    try:
        arg = argv[1]
        arg1 = argv[2]
    except:
        com('-h','')
    com(arg,arg1)
