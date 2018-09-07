# -*- coding: utf-8 -*-
# !/usr/bin/python
# Create Date 2018/8/23 0023
__author__ = 'huohuo'
import os
import sys
from jy_word.File import File

base_dir = os.path.dirname(__file__)
data_dir_name = '100303v3'

img_info_path = 'static/base_data/img_info%s.json' % data_dir_name
data_dir = os.path.join(base_dir, data_dir_name)
if len(sys.argv) > 1:
    data_dir = sys.argv[1]
images_dir = os.path.join(base_dir, 'images')
pdf_path = os.path.join(base_dir, 'images', 'pdf')
my_file = File(base_dir)
patient_info = my_file.read('patient_info.tsv', dict_name=data_dir)[0]
disease_name = patient_info['diagnose']
variant_knowledge_names = ['结直肠癌', '非小细胞肺癌', '骨肉瘤除外硬纤维瘤和肌纤维母细胞瘤']
variant_knowledge_name = '化疗多态位点证据列表%s' % variant_knowledge_names[0]
for v in variant_knowledge_names:
    if disease_name in v:
        variant_knowledge_name = '化疗多态位点证据列表%s' % v
        break
print disease_name, variant_knowledge_name, data_dir
# print '\t'.join('CRLF2, JAK1, JAK2, JAK3, SOCS1, STAT1, STAT2, STAT3, STAT4, STAT5A, STAT5B, STAT6'.split(', '))

if __name__ == "__main__":
    pass
    

