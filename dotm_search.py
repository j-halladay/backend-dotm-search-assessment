#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "j_halladay"
import argparse
import os
from glob import glob, iglob
import zipfile
def main():
    occur = 0
    files = 0
    parser = argparse.ArgumentParser(description='Find string inside DOTM files')
    parser.add_argument('string', type=str,
                        help='search criteria')
    parser.add_argument("-d", '--dir', nargs="?" , const=str, default=os.getcwd(),
                        help='specifying the directory')
    args = parser.parse_args()
    search_directory = args.dir
    for fn in iglob(search_directory+"/*.dotm"):
       
        document = zipfile.ZipFile(fn)
        read_doc = document.read("word/document.xml", pwd=None)
        stopper = 0
        for line in read_doc:
            if args.string in line:
                if stopper != 1:
                    occur += 1
                    stopper = 1
                print("Match found in file " + os.path.abspath(fn) + "\n" + read_doc[read_doc.index(line)-40:read_doc.index(line)+40])
        files+=1
    print("dot_m files searched: " + str(files)+"\n"+"dot_m matches found: "+str(occur))
if __name__ == '__main__':
    main()
