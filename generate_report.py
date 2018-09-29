# -*- coding:UTF-8 -*-
import os
import sys
from report_doc.report_core import get_report_core, patient_info, disease_name, variant_knowledge_names, variant_knowledge_index, out_path
from jy_word.File import File
from config import base_dir, data_dir_name
reload(sys)
sys.setdefaultencoding('utf-8')
result_dir = os.path.join(base_dir, 'results')
my_file = File(base_dir)

if __name__ == '__main__':
    report_title_cn = u'多组学临床检测报告'
    report_title_en = 'AIomics1'
    user_name = patient_info['name']

    file_name = u'results/%s_%s_%s(%s)检测报告.doc' % (data_dir_name, user_name, disease_name, variant_knowledge_names[variant_knowledge_index])
    # os.system('start explorer D:')
    if out_path != '':
        file_name = out_path
    print u'%s begin.' % file_name
    data = {'patient_info': patient_info}
    pkg = get_report_core(report_title_cn, report_title_en, data)
    # if variant_knowledge_index == 2:
    my_file.download(pkg, file_name)
    # os.startfile(os.path.abspath(file_name))
    # os.close()
    # os.system('start explorer %s' % os.path.abspath('results'))
    print u'%s over.' % file_name
