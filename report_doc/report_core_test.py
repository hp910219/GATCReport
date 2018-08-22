#! /usr/bin/env python
# coding: utf-8
__author__ = 'huo'
import math
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

from report_doc.tools.File import base_dir
from report_doc.tools.Word import Paragraph, Run, Set_page, Table, Tc, Tr, HyperLink
from report_doc.tools.Word import write_pkg_parts, write_cat, get_img_info
from report_doc.report_data import *

my_file = File()

r = Run()
hyperlink = HyperLink()
r.family_en = 'Times New Roman'
p = Paragraph()
set_page = Set_page().set_page
table = Table()
tr = Tr()
tc = Tc()

sect_pr_catalog = set_page('A4', footer='rIdFooter1', header='rIdHeader1')
sect_pr_content = set_page('A4', footer='rIdFooter2', pgNumType_s=1, header='rIdHeader2')
con1 = p.write(p.set(rule='exact', line=12, sect_pr=set_page(type='continuous', cols=1)))
con2 = p.write(p.set(rule='exact', line=12, sect_pr=set_page(type='continuous', cols=2, space=40)))
con3 = p.write(p.set(rule='exact', line=12, sect_pr=set_page(type='continuous', cols=3, space=40)))
page_br = p.write(r.br('page'))

# 初号=42磅
# 小初=36磅
# 一号=26磅
# 小一=24磅
# 二号=22磅
# 小二=18磅
# 三号=16磅
# 小三=15磅
# 四号=14磅
# 小四=12磅
# 五号=10.5磅
# 小五=9磅
# 六号=7.5磅
# 小六=6.5磅
# 七号=5.5磅
# 八号=5磅


# ##################下载报告所需方法######################
def get_report_core(title_cn, title_en, data):
    img_info = []
    body = write_body(title_cn, title_en, data)
    titles = []
    pkgs1 = write_pkg_parts(img_info, body, title=titles)
    return pkgs1


def write_body(title_cn, title_en, data):
    chem_items, trs3 = get_data3(chem_durg_list)
    diagnose = data['patient_info']['diagnose']
    data['target_tips'] = get_target_tips(diagnose)
    data['chem_tip'] = ', '.join(trs3[1][1])
    body = write_patient_info(data['patient_info'])
    return body


def write_patient_info(data):
    para = p.h4('基本信息')
    trs = ''
    ws = [2300, 2300, 3300, 2300]
    pPr = p.set(jc='left', spacing=[1, 1])
    ps = [
        '姓名: %s' % data['name'],
        '性别: %s' % data['sex'],
        '年龄: %s' % data['age'],
        '患者ID: %s' % data['id'],
        '医院: %s' % data['hospital'],
        '病理诊断: %s' % data['diagnose'],
        '组织类型: %s' % data['tissue_type'],
        '组织来源: %s' % data['tissue_source']
    ]
    n = len(ps) / 2
    for k in range(2):
        tr2 = ps[k * n: (k+1) * n]
        tcs2 = ''
        size = 10
        for i in range(len(tr2)):
            ts = tr2[i].split(':')
            run = r.text(ts[0].strip() + ': ', size=size, weight=0)
            run += r.text(ts[1].strip(), size=size, weight=1)
            tcs2 += tc.write(p.write(pPr, run), tc.set(w=ws[i], color=red, fill=gray, tcBorders=['bottom', 'right']))
        trs += tr.write(tcs2)
    # para += table.write(trs, ws=ws, bdColor=gray, insideColor=white)
    para += table.write(trs, ws=ws, bdColor=gray)
    return para
