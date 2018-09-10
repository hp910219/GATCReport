# -*- coding: utf-8 -*-
# !/usr/bin/python
# Create Date 2018/8/23 0023
__author__ = 'huohuo'
import os
import sys
from jy_word.File import File

base_dir = os.path.dirname(__file__)
data_dir_name = '100304v3'
data_dir = data_dir_name
if len(sys.argv) > 1:
    data_dir = sys.argv[1]
    data_dir_name = data_dir.split('/')[-1]
img_info_path = os.path.join(data_dir, 'img_info.json')
images_dir = os.path.join(base_dir, 'images')
pdf_path = os.path.join(base_dir, 'images', 'pdf')


if __name__ == "__main__":
    pass
    

