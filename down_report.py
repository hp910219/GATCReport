# -*- coding:UTF-8 -*-
import os
import sys
from flask import Flask, send_from_directory
from report_doc.report_core import get_report_core
from report_doc.tools.File import File, base_dir
reload(sys)
sys.setdefaultencoding('utf-8')
out_file = 'TCRseq_results.xlsx'
dir_name = base_dir
result_dir = os.path.join(dir_name, 'results')

app = Flask(__name__)
my_file = File()


@app.route('/')
def download_docx():
    patient_infos = my_file.read('0.1.1.patient_info.tsv', dict_name='data')
    patient_info = patient_infos[0]
    user_name = patient_info['name']
    file_name = u'results/%s检测报告.doc' % user_name
    print u'%s begin.' % file_name
    data = {'patient_info': patient_info}
    pkg = get_report_core(report_title_cn, report_title_en, data)
    my_file.download(pkg, file_name)
    dict_url = os.path.join(dir_name, file_name.split('/')[0])
    print u'%s over.' % file_name
    return send_from_directory(dict_url, file_name.split('/')[1], as_attachment=True)


if __name__ == '__main__':
    report_title_cn = u'多组学临床检测报告'
    report_title_en = 'AIomics1'
    app.run(port=1478, debug=False)
