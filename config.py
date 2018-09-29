# -*- coding: utf-8 -*-
# !/usr/bin/python
# Create Date 2018/8/23 0023
__author__ = 'huohuo'
import os
import sys

base_dir = os.path.dirname(__file__)
data_dir = r'C:\Users\guhongjie\Downloads\100307\100307'
data_dir = r'test000'
if len(sys.argv) > 1:
    data_dir = sys.argv[1]
data_dir_name = os.path.relpath(data_dir, os.path.dirname(data_dir))
img_info_path = os.path.join(data_dir, 'img_info.json')
images_dir = os.path.join(base_dir, 'images')
pdf_path = os.path.join(base_dir, 'images', 'pdf')
# for i in os.path
variant_dir = ''
for i in os.listdir(data_dir):
    path_file = os.path.join(data_dir, i)
    if i.startswith('variant_'):
        variant_dir = path_file
        break
if variant_dir == '':
    print 'check variant folder please'
else:
    print 'variant folder', variant_dir

if __name__ == "__main__":
    pass
    

