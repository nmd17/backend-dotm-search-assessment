#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "nmd17"



import argparse
import os
import zipfile
try:
    from xml.etree.cElementTree import XML
except ImportError:
    from xml.etree.ElementTree import XML

def create_parser():
    parser = argparse.ArgumentParser()
    parser.add_argument("-d", "--dir", help="directory name")
    parser.add_argument('text', help="text to search for within a .dotm file")
    
    return parser

def get_all_file_paths(directory): 
    file_paths = [] 

    for root, directories, files in os.walk(directory): 
        for filename in files: 
         
            filepath = os.path.join(root, filename) 
            file_paths.append(filepath) 
  
    return file_paths

def search_zipfile(file_name, text):
        document = zipfile.ZipFile(file_name)
        matches_found = 0

        with document.open('word/document.xml') as xml_doc:
            for line in xml_doc:
                temp = line.find(text)
                while temp != -1:
                    print(line[temp - 40: temp + 40])
                    temp = line.find('$', temp + 1)
                    matches_found += 1
           
        return matches_found

def main():
    parser = create_parser()
    args = parser.parse_args()
    file_paths = get_all_file_paths('./dotm_files')

    total_matches = 0
    file_total = 0

    for file in file_paths:
        if file.endswith('.dotm'):
            file_total += 1
            matches = search_zipfile(file, '$')
            if matches > 0:
                print('There was a match/matches found in this file:  ')
                total_matches += 1       
        else:
            continue
        

    print('The total number of matches was: ' + str(total_matches), ' The total number of files searched was: ' + str(file_total))


        

                    
if __name__ == '__main__':
    main()
