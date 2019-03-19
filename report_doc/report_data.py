#! /usr/bin/env python
# coding: utf-8
__author__ = 'huo'
import os
import time
import sys
from PIL import Image
reload(sys)
sys.setdefaultencoding('utf-8')
from jy_word.Word import bm_index0, get_imgs, uniq_list, get_img
from jy_word.File import File
from config import img_info_path, base_dir, data_dir, images_dir, variant_dir
from report_doc.get_real_level import FetchRealLevel

gray = 'E9E9E9'
gray_lighter = 'EEEEEE'
blue = '00ADEF'
white = 'FFFFFF'
red = 'ED1C24'
orange = 'F14623'
colors = ['2C3792', '3871C1', '50ADE5', '37AB9C', '27963C', '40B93C', '80CC28']
level_names = ['A', 'B', 'C', 'D']
db_names = ['OncoKB', 'Civic', 'CGI']
borders = ['top', 'right', 'bottom', 'left']
# 因为数据库的权威性  oncokb>civic>cgi

real_level = FetchRealLevel()
my_file = File(base_dir)

# 初始化
patient_info = my_file.read('summary/patient_info.tsv', dict_name=data_dir)[0]
disease_name = patient_info['diagnose']
evidence_dir = os.path.join(data_dir, 'evidence')
variant_knowledge_names = [u'结直肠癌', u'非小细胞肺癌', u'肉瘤', u'胃癌', u'胰腺癌', u'乳腺癌', u'肝癌']
variant_knowledge_index = 0
out_path = ''
for indexx, v in enumerate(variant_knowledge_names):
    if disease_name in v:
        variant_knowledge_index = indexx
        break
var_sys_arg = ''
if len(sys.argv) > 2:
    if sys.argv[2].endswith('.doc'):
        out_path = sys.argv[2]
    else:
        var_sys_arg = sys.argv[2]
    if len(sys.argv) > 3:
        var_sys_arg = sys.argv[3]
var_sys_arg = var_sys_arg.decode(sys.stdin.encoding).encode('utf-8')
if var_sys_arg in variant_knowledge_names:
    variant_knowledge_index = variant_knowledge_names.index(var_sys_arg)
try:
    variant_knowledge_index = int(var_sys_arg)
except:
    variant_knowledge_index = -1
var_know_name = variant_knowledge_names[variant_knowledge_index]
variant_knowledge_name = u'化疗多态位点证据列表%s' % var_know_name
tmb_tip = '注：NSCLC未经选择人群PD抗体有效率，具吸烟史为22%，无吸烟史为10%'
if disease_name in '结直肠癌':
    tmb_tip = '注：MSS微卫星稳定结直肠癌患者PD1抗体有效率为0%；MSI-H微卫星不稳定结直肠癌患者有效率为29.6%。'
print u'%s' % ','.join(variant_knowledge_names)
print u'%s' % disease_name, variant_knowledge_name, data_dir
# 静态的数据：
variant_knowledge = my_file.read(u'static/base_data/%s.xlsx' % variant_knowledge_name, sheet_name='Sheet1')
gene_list12 = my_file.read('static/base_data/1.2gene_list.json')
gene_list53 = my_file.read('static/base_data/5.3gene_list.xlsx', sheet_name='Sheet2')
company = my_file.read('static/base_data/5.company.txt', sheet_name='Sheet2')
gene_MoA = my_file.read('static/base_data/gene_MoA.tsv')
signature_cn = my_file.read('static/base_data/signature_cn.txt')
evidence4 = []
for evidence4_index in range(5):
    evidence_path = 'static/base_data/4.%d.evidence.txt' % evidence4_index
    if os.path.exists(evidence_path):
        evidence = my_file.read(evidence_path, dict_name='').split('\n')
    else:
        evidence = []
    evidence4.append(evidence)

# 动态的数据：
immune_suggestion = my_file.read(r'summary/immune_suggestion.tsv', dict_name=data_dir, to_json=False)
neoantigen = my_file.read('neoantigen.tsv', dict_name=variant_dir)
qc = my_file.read('QC.tsv', dict_name=variant_dir)
quantum_cellurity = my_file.read('quantum_cellurity.tsv', dict_name=variant_dir)
rs_geno = my_file.read('rs.geno.update.tsv', to_json=False, dict_name=variant_dir)[1:]
variant_anno1 = my_file.read('variant_anno.maf', dict_name=variant_dir, sep='\t', to_json=False)
cnv_copynumber = my_file.read('cnv/cnv.copynumber.table.tsv', dict_name=data_dir)
tumor_bam_info = my_file.read('tumor.recal.bam_info.txt', dict_name=variant_dir)
tumor_bam_CNVs = my_file.read('tumor.recal.bam_CNVs', dict_name=variant_dir, sep='\t')
CGI_mutation_analysis = my_file.read('CGI_mutation_analysis.tsv', dict_name=data_dir)


def test_chinese(contents):
    import re
    zhPattern = re.compile(u'[\u4e00-\u9fa5]+')
    # 一个小应用，判断一段文本中是否包含简体中：
    match = zhPattern.search(contents)
    if match:
        a = zhPattern.findall(contents)
        ch_num = 0
        for i in a:
            ch_num += len(i)
        return ch_num
    return 0


def transfer_level(level):
    # 1→FDA级别、2A→NCCN指南级别、3A→临床证据级别、4→生物学证据级别；R1→NCCN指南级别
    if level == '1':
        return 'FDA级别'
    if level == '2A':
        return 'NCCN指南级别'
    if level == '3A':
        return '临床证据级别'
    if level == '4':
        return '生物学证据级别'
    if level == 'R1':
        return 'NCCN指南级别'
    return '未知级别'


def filter_db(gene=None, var=None, impact=None):
    variant_annos = variant_anno1
    items = []
    keys = variant_annos[0]
    # print keys
    # print len(keys), keys
    for i in range(1, len(variant_annos)):
        line = variant_annos[i]
        if len(line) > 1:
            gene_db = line[0]
            var_db = line[36]
            impact_db = line[keys.index('IMPACT')]
            is_match = gene == gene_db
            match_tip = ''
            if var is not None:
                is_match = is_match and var_db == var
                if not is_match and var.strip() != '' and var_db.strip() != '':
                    if var.lower() in var_db.lower() or var_db.lower() in var.lower():
                        is_match = True
                        match_tip = '模糊匹配， %s , %s' % (var, var_db)
            if impact is not None:
                is_match = impact == impact_db
            if is_match:
                info = {'match_tip': match_tip}
                for j in range(len(keys)):
                    key = keys[j]
                    value = line[j]
                    info[key] = value
                items.append(info)
    return items


def get_catalog():
    catalogue = [
        [u"一、靶向治疗提示", 0, 1, 10],
        [u"（一）、驱动突变靶向治疗提示汇总", 2, 1, 23],
        [u"（二）、各驱动突变循证医学证据", 2, 1, 23],
        [u"（三）、靶向治疗解读说明", 2, 1, 23],
        [u"二、免疫治疗提示", 0, 1, 10],
        [u"（一）、MSI微卫星不稳定检测结果", 2, 1, 23],
        [u"（二）、TMB肿瘤突变负荷检测结果", 2, 2, 23],
        [u"（三）、肿瘤新抗原检测结果", 2, 2, 23],
        [u"（四）、免疫治疗解读说明", 2, 3, 23],
        [u"三、化学治疗提示", 0, 3, 10],
        [u"（一）、化学治疗提示汇总", 2, 4, 23],
        [u"（二）、各化疗药循证医学证据", 2, 5, 23],
        [u"四、最新研究进展治疗提示", 0, 5, 10],
        [u"（一）、DDR基因DNA损伤修复反应基因突变检测结果", 2, 6, 23],
        [u"（二）、HLA分型检测结果", 2, 6, 23],
        [u"（三）、免疫治疗超进展相关基因检测结果", 2, 7, 23],
        [u"（四）、免疫治疗耐药相关基因检测结果", 2, 7, 23],
        [u"（五）、肿瘤突变模式检测结果", 2, 8, 23],
        [u"（六）、肿瘤遗传性检测结果", 2, 8, 23],
        [u"（七）、最新研究进展解读说明", 2, 8, 23],
        [u"五、检测信息汇总", 0, 5, 10],
        [u"（一）、基因突变信息汇总", 2, 6, 23],
        [u"（二）、基因拷贝数信息汇总", 2, 6, 23],
        [u"（三）、肿瘤重要信号通路变异信息汇总", 2, 7, 23],
        [u"（四）、检测方法说明", 2, 7, 23]
    ]
    items = []
    for index, cat in enumerate(catalogue):
        item = {"title": cat[0], 'left': cat[1], 'page': cat[2], 'style': cat[3], 'bm': bm_index0 + index}
        items.append(item)
    return items


def format_time(t=None, frm="%Y-%m-%d %H:%M:%S"):
    if t is None:
        t = time.localtime()
    if type(t) == int:
        t = time.localtime(t)
    my_time = time.strftime(frm, t)
    return my_time


def get_page_titles():
    title_cn, title_en = u'多组学临床检测报告', 'AIomics1'
    title = '%s附录目录%s%s' % (title_cn, ' ' * 64,  title_en)
    title1 = '%s%s%s' % (title_cn, ' ' * (64+11),  title_en)
    titles = [title, title1]
    cats = get_catalog()
    for i in [0, 4, 9, 12, 20]:
        n = 64 + 14
        if test_chinese(cats[i]['title']) == 11:
            n -= 8
        titles.append('%s%s%s' % (cats[i]['title'], ' ' * n,  title_en))
    return titles


def get_img_info(path, is_refresh=False):
    crop_img(r'%s/signature.png' % (data_dir), r'%s/4.5.1signature.png' % data_dir)
    crop_img(r'%s/signature_pie.png' % (data_dir), r'%s/4.5.2signature_pie.png' % data_dir)
    if is_refresh:
        img_info = get_imgs(images_dir)
        img_info += get_imgs(path)
        img_info2 = []
        for info in img_info:
            is_exists = len(filter(lambda x: x['rId'] == info['rId'], img_info2))
            if is_exists == 0:
                img_info2.append(info)
        my_file.write(img_info_path, img_info2)
    else:
        img_info2 = my_file.read(img_info_path)
    return img_info2


def get_immu(source):
    for ims in immune_suggestion:
        if ims[0] == source:
            if len(ims) == 3:
                return ims
            elif len(ims) == 2:
                tr1s = ims[1].split('，')
                tr2 = tr1s[-1].strip()
                tr1 = ', '.join(tr1s[:-1])
                return [source, tr1, tr2]
    return [source, '', '']


# 选出部分重要抗原信息
def get_neoantigen():
    items1 = []
    for item in neoantigen:
        ref = item['RefAFF'].strip().strip('\n')
        try:
            item['MutRank'] = float(item['MutRank'])
        except:
            pass
        try:
            ref = float(ref)
            if ref > 500:
                items1.append(item)
        except Exception, e:
            pass
    items1.sort(key=lambda x: x['MutRank'])
    return items1


def get_var(gene):
    items = filter_db(gene)
    for item in items:
        impact = item['IMPACT'].strip().upper()
        if impact == 'HIGH':
            return red
        elif impact == 'MODERATE':
            return orange
    return gray


def get_EGFR():
    # efgr 只有突变， 只有扩增时， 都无时，合并？
    gene = 'EGFR'
    item1 = {'gene': gene, 'fill': gray, 'text': 'EGFR无突变无扩增', 'copynumber': 0}
    items = filter_db(gene)
    is_match, copynumber, status = get_copynumber(gene, 'gain')
    for item in items:
        impact = item['IMPACT'].strip().upper()
        variant_type = item['Variant_Type']
        if impact == 'HIGH':
            if reset_status(variant_type) == 'gain':
                if is_match:
                    return {'gene': gene, 'fill': red, 'text': 'EGFRT突变合并扩增', 'copynumber': copynumber}
                return {'gene': gene, 'fill': gray, 'text': 'EGFR无扩增', 'copynumber': copynumber}
        if is_match:
            return {'gene': gene, 'fill': red, 'text': 'EGFR基因扩增', 'copynumber': copynumber}
    return item1


def get_11q13():
    # 11q13标红：仅当CCND1，FGF3，FGF4，FGF19全部扩增（GAIN） 拷贝数？
    genes = ['CCND1', 'FGF3', 'FGF4', 'FGF19']
    title = '11q13(%s)' % ('、'.join(genes))
    gain = []
    for gene in genes:
        is_match, copynumber, status = get_copynumber(gene, 'gain')
        if status == '扩增':
            gain.append('%s(%s)' % (gene, copynumber))
    if len(gain) == len(genes):
        text = '11q13(%s)扩增' % ('、'.join(gain))
        return {'gene': title, 'fill': red, 'text': text, 'copynumber': text}
    return {'gene': title, 'fill': gray, 'text': '%s未扩增' % title, 'copynumber': 0}


def get_target_tips(diagnose):
    # 不同的库
    # T790M突变在oncokb和civic和CGI数据库中，对于同样的肿瘤，同样的药物，其证据等级分别是C,A,B.  那么证据等级取C，
    # 因为数据库的权威性  oncokb>civic>cgi
    items = []
    items2 = []
    # ['OncoKB', 'CIVic', 'cgi']
    for i in os.listdir(evidence_dir):
        if i.lower().startswith('oncokb'):
            get_target_tip(items, items2, diagnose, i, reset_oncokb)
        if i.lower().startswith('civic'):
            get_target_tip(items, items2, diagnose, i, reset_civic)
        if i.lower().startswith('cgi'):
            get_target_tip(items, items2, diagnose, i, reset_cgi)
    vars = []
    for item2 in items2:
        var = {'col1': item2}
        for key in db_names + level_names:
            var[key.upper()] = []
        for item1 in items:
            if item1['col1'] == item2:
                var['gene'] = item1['gene']
                var['gene_MoA'] = item1['gene_MoA']
                var['gene_MoA1'] = item1['gene_MoA1']
                for key in db_names:
                    key = key.upper()
                    if key in item1:
                        value1 = item1[key]
                        if value1 not in var[key]:
                            var[key].append(value1)
                for key in level_names:
                    var[key] += item1[key]
        vars.append(var)
    extra = []
    extra_col1 = ''
    if diagnose == '结直肠癌':
        extra_gene = ['KRAS', 'NRAS']
        extra_col1 = '%s未发生突变' % '、'.join(extra_gene)
    elif diagnose in ['胃癌', '乳腺癌']:
        extra_gene = ['HER2']
        extra_col1 = '%s未发生扩增' % '、'.join(extra_gene)
    elif diagnose == '非小细胞肺癌':
        extra_gene = ['EGFR']
        extra_col1 = '%s未发生突变' % '、'.join(extra_gene)
    else:
        extra_gene = []
    show_extra = False
    vars2 = []
    for index, v in enumerate(vars):
        gene = v['gene']
        v2s = filter(lambda x: x['gene'] == gene and '模糊匹配' in x['col1'], vars)
        for v2 in v2s:
            for k in level_names:
                v[k] += v2[k]
            for k in db_names:
                v[k.upper()] += v2[k.upper()]
            if v2 in vars:
                vars.remove(v2)
        v = get_drug_db(v)
        if v is not None:
            vars2.append(v)
            if v['gene'] in extra_gene:
                if 'EGFR' in extra_gene:
                    if '亚克隆' in v['col1']:
                        extra.append(v)
                elif 'HER2' in extra_gene:
                    if '扩增' in v['col1']:
                        extra.append(v)
                else:
                    extra.append(v)
    if len(extra) == 0 and len(extra_gene) > 0:
        show_extra = True
    extra_item = {'col1': extra_col1, 'gene_MoA': ''}
    for k in level_names:
        extra_item[k] = ''
    vars2 = sorted(vars2, cmp=cmp_target_tip)
    return vars2, show_extra, extra_item


def cmp_target_tip(x, y):
    cmp1 = -1
    if len(x['A']) > len(y['A']):
        return cmp1
    if len(x['A']) == len(y['A']):
        if len(x['B']) > len(y['B']):
            return cmp1
        if len(x['B']) == len(y['B']):
            if len(x['C']) > len(y['C']):
                return cmp1
            if len(x['C']) == len(y['C']):
                if len(x['D']) > len(y['D']):
                    return cmp1
    return cmp1 * (-1)


def get_target_tip(items, items2, diagnose, file_name, func):
    db_items = my_file.read(os.path.join(evidence_dir, file_name))
    i = 0
    for db_item in db_items:
        # 各数据库证据级别对应关系 https://mubu.com/doc/3LSWugo_cE
        # [数据库名, level, 肿瘤类型, 标准化肿瘤类型(汉语)]
        # level( "cgi", "Late trials" , "NSCLC", "肺癌")
        # 癌种是L列。当前肿瘤就是前面基本信息里获得的。
        # level(  "civic", "D: Preclinical evidence", "Ewing Sarcoma", "肺癌")
        # @霍  癌种是diseasename 列
        # level ( “oncokb”,  “R1”， “Non-Small Cell Lung Cancer”，  "肺癌")
        reset_item = func(db_item)
        is_match = False
        if reset_item is not None:
            gene = reset_item['gene']
            i += 1
            for v in reset_item['vars']:
                v = 'p.' + v
                # if [gene, v] not in genes:
                #     genes.append([gene, v])
                if len(gene) > 0:
                    for maf_item in filter_db(gene, v):
                        is_match = True
                        cellurity = get_quantum_cellurity(maf_item['Chromosome'], maf_item['Start_Position'])
                        col1 = '%s %s\n%s\n(%s肿瘤亚克隆)' % (gene, get_exon(maf_item), maf_item['HGVSp_Short'], cellurity)
                        item = get_drug(col1, reset_item, diagnose)
                        item['var'] = v
                        items.append(item)
                        if col1 not in items2:
                            items2.append(col1)
            exon1 = reset_item['exon']
            variant_type1 = reset_item['status']
            if len(reset_item['vars']) == 0:
                col1 = ''
                if exon1 != '' and variant_type1 != '':
                    items1 = filter_db(gene)
                    if len(items1) > 0:
                        info = items1[0]
                        exon = get_exon(info)
                        variant_type = info['Variant_Type']
                        col1 = ''
                        # print exon1, variant_type1
                        if exon1 == exon and variant_type1[:3].upper() == variant_type:
                            is_match = True
                            col1 = '%s exon %s %s' % (gene, exon, reset_item['status'])
                elif exon1 == '' and variant_type1 != '':
                    is_match, copynumber, status = get_copynumber(gene, variant_type1)
                    # if gene == 'FBXW7':
                    #     print is_match, copynumber, status
                    col1 = u'%s%s(拷贝数%s)' % (gene, status, copynumber)
                if len(col1) > 0 and is_match:
                    item = get_drug(col1, reset_item, diagnose)
                    item['var'] = ''
                    items.append(item)
                    if col1 not in items2:
                        items2.append(col1)
            if is_match is False:
                col1 = u'%s(模糊匹配)' % gene
                reset_item['evidence_detail'][4] = '模糊匹配'
                item = get_drug(col1, reset_item, diagnose)
                item['var'] = ''
                items.append(item)
                if col1 not in items2:
                    items2.append(col1)
    return items, items2


def get_drug(col1, reset_item, diagnose):
    gene = reset_item['gene']
    db_name = reset_item['db_name']
    drug = reset_item['drug']
    col6 = get_gene_MoA(gene)  # 原癌都是激活，抑癌都是失活。
    item = {'col1': col1, 'gene': gene, 'gene_MoA': ''.join(col6), 'gene_MoA1': col6[0], 'db_name': db_name}
    item['responsive'] = reset_item['responsive'].lower().strip()
    item[db_name.upper()] = reset_item['evidence_detail']
    this_level = real_level.level(db_name.lower(), reset_item['level'], reset_item['cancer_type'], diagnose)
    item[db_name.upper()].append(this_level)
    # item[2] drug
    drug_items = []
    drug_name = drug.strip().split('(')[0].strip()
    for x1 in drug_name.split(','):
        xx1 = x1.strip()
        if xx1.lower() not in ['', 'na', '5-fluorouracil', 'platinum']:
            x2 = {'drug': xx1, 'db_name': db_name, 'level': this_level, 'responsive': reset_item['responsive']}
            if x2 not in drug_items:
                drug_items.append(x2)
    item[db_name.upper()][2] = drug_items
    for level in level_names:
        item[level] = [] if level != this_level else drug_items
    return item


def get_drug_level(drug_items):
    for i in range(len(level_names)):
        level1 = level_names[i]
        drugs1 = filter(lambda x: x['level'] == level1, drug_items)
        for j in range(i + 1, len(level_names)):
            level2 = level_names[j]
            for drug in drugs1:
                drugs2 = filter(lambda x: x['drug'] == drug['drug'] and x['level'] == level2, drug_items)
                for drug2 in drugs2:
                    drug_items.remove(drug2)
    return drug_items


def get_drug_db(var):
    # oncokb>civic>cgi
    drug_items = []
    for level_name in level_names:
        for v in var[level_name]:
            drug_items.append(v)
    for i in range(len(db_names)):
        db_name1 = db_names[i]
        drug_items1 = filter(lambda x: x['db_name'].lower() == db_name1.lower(), drug_items)
        for j in range(i+1, len(db_names)):
            db_name2 = db_names[j]
            drug_items2 = filter(lambda x: x['db_name'].lower() == db_name2.lower(), drug_items)
            for drug_item1 in drug_items1:
                for drug_item2 in drug_items2:
                    if drug_item2['drug'] == drug_item1['drug']:
                        drug_item2['db_name'] = db_name1
                        drug_item2['level'] = drug_item1['level']
                        if drug_item2['responsive'] != drug_item1['responsive']:
                            drug_item1['color'] = 'red'
                            drug_item2['color'] = 'red'
    drug_items = get_drug_level(drug_items)
    drug_items = get_drug_level(drug_items)
    if len(drug_items) == 0:
        var = None
        return var
    for level in level_names:
        value1 = []
        value2 = []
        drugs = filter(lambda x: x['level'] == level, drug_items)
        for drug in drugs:
            responsive = drug['responsive'].lower()
            x_drug = '%s' % drug['drug'].strip()
            # x_drug = '%s==%s' % (drug['drug'].strip(), responsive)
            if x_drug.lower() not in ['', 'na', '5-fluorouracil', 'platinum']:
                x_drug1 = '%s(%s)' % (x_drug, responsive)
                if 'color' in drug:
                    x_drug1 += '==%s' % drug['color']
                if responsive == 'responsive':
                    if x_drug1 not in value1:
                        value1.append(x_drug1)
                else:
                    if x_drug1 not in value2:
                        value2.append(x_drug1)
        var[level] = ';'.join(sorted(value1) + sorted(value2))
    return var


def get_exon(item):
    if 'EXON' in item:
        exon = item['EXON']
        if type(exon) in [unicode, str]:
            if '/' in exon:
                return 'E%s' % (exon.split('/')[0])
    return ''


def get_copynumber(gene, status=None):
    for item in cnv_copynumber:
        if gene in item['Symbol']:
            if status is None:
                return item['status'], item['copy.number']
            ss = reset_status(status)
            if item['status'] == ss[0]:
                return True, item['copy.number'], ss[1]
    if status is None:
        return False, 0
    return False, 0, status


def get_quantum_cellurity(chr, pos, is_print=False):
    for item in quantum_cellurity:
        if isinstance(chr, float):
            chr = str(int(chr))
        if isinstance(pos, float):
            pos = str(int(pos))
        # print chr, pos
        if is_print:
            print item['Start'], pos, item['Chr'].lstrip('chr'), chr, type(pos), type(item['Start']), type(chr)
        if item['Chr'].lstrip('chr') == str(chr) and str(item['Start']) == str(pos):
            f = float(item['Cellularity'])
            if f > 1:
                f = 1.0
            return '%.2f%%' % (f * 100)
    return ''


def get_data3():
    new_items = []
    trs = [
        ['', '疗效可能好', '疗效可能差', '疗效未知'],
        ['毒副作用低', [], [], []],
        ['毒副作用高', [], [], []],
        ['毒副作用未知', [], [], []]
    ]
    drugs = get_variant_knowledges()
    for item in drugs:
        new_item = {}
        category = item['category']
        new_item['category'] = category
        new_item['drug'] = item['drug']
        new_item['genes'] = item['genes']
        new_item = get_data31(new_item)
        cell = new_item['cell']
        cell_value = '%s%s' % (item['category'], new_item['drug'])
        row, col = cell[0], cell[1]
        if row * col > 0:
            trs[row][col].append(cell_value)
        new_items.append(new_item)
    return new_items, trs


def get_data31(item):
    venoms = [[], [], [], []]
    curative_effects = [[], [], [], []]
    liaoxiao = 0
    duxing = 0
    for item1 in item['genes']:
        summary = item1['summary'].split('、')
        venom = summary[0]  # 毒副作用
        curative_effect = '疗效未知'
        if len(summary) > 1:
            curative_effect = summary[1]
        else:
            if summary[0].startswith(u'疗效'):
                curative_effect = summary[0]
                venom = '毒副作用未知'
        if '好' in curative_effect:
            curative_effects[1].append(item1)
        elif '差' in curative_effect:
            curative_effects[2].append(item1)
        # else:
        #     curative_effects[3].append(item1)
        if '疗效' in item1['summary']:
            liaoxiao += 1
        if '毒' in item1['summary']:
            duxing += 1
        if '低' in venom:
            venoms[1].append(item1)
        elif '高' in venom:
            venoms[2].append(item1)
        # else:
        #     venoms[3].append(item1)

    venoms_num = [len(x) for x in venoms]
    curative_effects_num = [len(x) for x in curative_effects]
    good = len(curative_effects[1])
    bad = len(curative_effects[2])
    low = len(venoms[1])
    high = len(venoms[2])
    row, col = venoms_num.index(max(venoms_num)), curative_effects_num.index(max(curative_effects_num))

    # 判断逻辑问题：
    # 根据证据的数量确定推荐的方向，这种情况下，证据的权重一致；
    # 当相反证据量数量一致时，根据证据级别确定（证据级别的等级 1A＞1B＞2A＞2B＞3）
    if good == bad:
        col = compare_level3(curative_effects[1], curative_effects[2])
    if low == high > 0:
        row = compare_level3(venoms[1], venoms[2])
    item['cell'] = row, col
    item['venoms'] = venoms
    item['curative_effects'] = curative_effects
    tr1 = '疗效预测方面共纳入%d个证据，其中%d个预测疗效好，%d个预测疗效差；' % (liaoxiao, good, bad)
    tr1 += '毒副作用预测共纳入%d个证据，其中%d个预测毒副作用低，%d个预测毒副作用高' % (duxing, low, high)
    item['tr1'] = tr1
    return item


def get_data43():
    # MDM2  查CNV表，看是否为GAIN
    # MDM4同MDM2
    # DNMT3A查maf表，看是否有突变
    # EGFR是查maf表和CNV表。
    # 11q13标红：仅当CCND1，FGF3，FGF4，FGF19全部扩增（GAIN）
    # 扩增=GAIN=amplication
    genes = ['MDM2', 'MDM4', 'DNMT3A', 'EGFR']
    items = []
    for gene in genes[:2]:
        fill, text = gray, '基因未扩增'
        status, copynumber = get_copynumber(gene)
        if status == 'gain':
            fill = red
            text = '基因扩增'
        items.append({'gene': gene, 'fill': fill, 'text': '%s%s' % (gene, text), 'copynumber': copynumber})
    fill3 = get_var(genes[2])
    text3 = '无突变' if fill3 == gray else '突变'
    items.append({'gene': genes[2], 'fill': fill3, 'text': '%s%s' % (genes[2], text3), 'copynumber': 0})
    items.append(get_EGFR())
    items.append(get_11q13())
    return items


def get_data44(gene):
    # 纯合失活突变：
    # 纯合：突变频率 >0.7  ，突变频率=AS/AQ
    # 失活：跟恶性突变是一样的，是指maf中的impact为HIGH
    items = filter_db(gene)
    for item in items:
        try:
            freq = float(item['n_alt_count']) / float(item['n_depth'])
        except:
            freq = 0
        if freq > 0.7 and item['IMPACT'] == 'HIGH':
            return red
    return gray


def get_data45(signature_etiology):
    items = [['', 'Signature ID', '权重', '主要出现的癌种', '可能的病因', '评论']]
    tr1, tr2 = '', []
    for item in signature_etiology:
        s_id = item['Signature_ID']
        item['Keyword'] = reset_sig(s_id)
        if s_id != 'unknown':
            result = '原因未知'
            if item['Keyword'] != result:
                result = '推测由%s导致' % item['Keyword']
                tr2.append(item['Keyword'])
            tr1 += '特征%s所占权重为%s，%s;' % (s_id, item['Weight'], result)
            items.append(['', s_id, item['Weight'], item['Cancer_types'], item['Proposed_aetiology'], item['Comments']])
    tip = '该癌症可能由%s' % ('、'.join(tr2))
    if len(tr2) > 1:
        tip += '共同'
    return uniq_list(items), tr1.rstrip(';'), tip + '导致'


def get_sample_purity():
    if tumor_bam_info is None:
        return ''
    texts = tumor_bam_info.split('\n')
    for t in texts:
        if t.lower().strip().startswith('sample_purity'):
            sample_purity = t[len('sample_purity'):].strip().strip('\t').strip()
            try:
                sample_purity = float(sample_purity)
            except:
                return sample_purity
            return '%.2f%%' % (sample_purity * 100)
    return ''


def get_bam_snvs():
    if tumor_bam_CNVs is None:
        return 0
    return len(tumor_bam_CNVs)


def get_qc():
    s = {'i': '', 'a': ''}
    if qc is not None:
        for k in s.keys():
            for item in qc:
                base_aligned = get_base_aligned(item, k)
                if base_aligned[k] != '':
                    s[k] = base_aligned[k]
    return s


def get_base_aligned(item, i):
    s = {}
    s[i] = ''
    if 'Sample_ID' in item:
        if item['Sample_ID'].strip().lower().startswith(i):
            if 'Base_Aligned' in item:
                Base_Aligned = item['Base_Aligned']
                try:
                    base_aligned = float(Base_Aligned)
                except:
                    s[i] = Base_Aligned
                    return s
                s[i] = '%.2f' % (base_aligned / (1024 * 1024 * 1024))
    return s


def get_gene_MoA(gene):
    # 原癌都是激活，抑癌都是失活
    # 共三类：原癌、抑癌、未知
    # Act : 原癌；Lof：抑癌， 其他： 未知
    for item in gene_MoA:
        if item['gene'] == gene:
            if item['gene_MoA'].lower() == 'Act'.lower():
                return ['原癌基因', '激活突变']
            if item['gene_MoA'].lower() == 'Lof'.lower():
                return ['抑癌基因', '失活变异']
    return ['未知', '']


def reset_status(status):
    for a in ['amplification', 'gain', 'amp']:
        if status.lower().startswith(a):
            return 'gain', '扩增'
    for b in ['loss', 'deletions', 'deletion']:
        if status.lower().startswith(b):
            return 'loss', '缺失'
    return status, '未知'


def reset_cgi(item):
    # 敏感（responsive）的药物标蓝；耐药（resistance）
    db_name = 'CGI'
    # 各数据库证据级别对应关系 https://mubu.com/doc/3LSWugo_cE
    # [数据库名, level, 肿瘤类型, 标准化肿瘤类型(汉语)]
    # level( "cgi", "Late trials" , "NSCLC", "肺癌")
    # 其中Biomarker对应生物标志物，Drug对应药物，Tested_tumor+Evidence+Effect 对应证据类型，Source对应证据来源，匹配程度全部为完全匹配，后续有特殊需要标注的再行手动处理。
    alteration = item['BIOMARKER']  # 生物标志物
    vars = []
    biomaker = alteration.strip('"').split(' ')
    gene = biomaker[0]
    if len(biomaker) > 1:
        a = biomaker[1].split('(')
        if len(a) > 1:
            vars = a[1].rstrip(')').split(',')
    drug = item['DRUG']
    exon = ''
    status = ''
    if len(vars) == 0:
        if len(biomaker) == 4:
            exon = biomaker[2]
            status = biomaker[3]
        elif len(biomaker) == 2:
            # D:\pythonproject\report\data\cnv\cnv.copynumber.table.tsv
            status = biomaker[1]
        sample_alteration = item['SAMPLE_ALTERATION'].split(',')
        for s in sample_alteration:
            ss = s.strip().split('MUT')
            if len(ss) > 1:
                var = ss[1].lstrip().lstrip('*').lstrip().lstrip('(').strip().rstrip(')')
                if var not in vars:
                    vars.append(var)
    cancer_type, evidence = item['TESTED_TUMOR'], item['EVIDENCE']
    evidence_type = ', '.join([cancer_type, evidence, item['EFFECT']])
    # 基因 生物标志物 药物 证据类型 匹配程度 证据来源
    # 敏感（responsive）的药物标蓝；耐药（resistance）
    if item['EFFECT'].lower() == 'responsive':
        responsive = 'responsive'
    elif item['EFFECT'].lower() in ['resistance', 'resistant', 'no responsive']:
        responsive = 'resistance'
    else:
        responsive = item['EFFECT']
    if alteration.lower().startswith(gene.lower()) is False:
        alteration = ' '.join([gene, alteration])
    return {
        'responsive': responsive,
        'gene': gene,
        'vars': vars,
        'db_name': db_name,
        'drug': drug,
        'level': evidence,
        'cancer_type': cancer_type,
        'exon': exon,
        'status': status,
        'evidence_detail': [db_name, alteration, drug, evidence_type, '完全匹配', item['SOURCE']],
    }


def reset_civic(item):
    db_name = 'civic'
    # 各数据库证据级别对应关系 https://mubu.com/doc/3LSWugo_cE
    # [数据库名, level, 肿瘤类型, 标准化肿瘤类型(汉语)]
    # level( "cgi", "Late trials" , "NSCLC", "肺癌")
    # 癌种是L列。当前肿瘤就是前面基本信息里获得的。
    # level(  "civic", "D: Preclinical evidence", "Ewing Sarcoma", "肺癌")
    # @霍  癌种是diseasename 列
    # level ( “oncokb”,  “R1”， “Non-Small Cell Lung Cancer”，  "肺癌")
    vars = []
    variant_name = item['variant_name']
    status_en, status_cn = reset_status(variant_name)
    exon = ''
    gene = item['genesymbol']
    drug = item['drug_names']
    cancer_type = item['disease_name']
    if status_cn == '未知':
        variant_names = variant_name.split(' ')
        if variant_name.startswith('Exon'):
            exon = 'E%s' % variant_names[1]
            status_en, status_cn = reset_status(variant_names[2])
        elif variant_name.startswith('Ex'):
            vars = [gene, variant_names[-1]]
            exon = 'E%s' % variant_names[0].lstrip('Ex')
            status_en, status_cn = reset_status(variant_names[1])
        else:
            # status_en, status_cn = reset_status(variant_name.split('-')[0])
            vars = [variant_name]
    evidence_level = item['evidence_level']
    evidence_type = ', '.join([item['disease_name'], evidence_level, item['clinical_significance'], item['evidence_direction'], item['rating']]) # 证据类型
    match_degree = '完全匹配'  # 匹配程度
    pmids = '%s(%s)' % (item['evidence_description'], item['pubmed_id'])  # 证据来源
    # 敏感（responsive）的药物标蓝；耐药（resistance）
    if item['clinical_significance'].lower() == 'sensitivity':
        responsive = 'responsive'
    elif item['clinical_significance'].lower().startswith('resistance') or item['clinical_significance'].lower().startswith('resistant'):
        responsive = 'resistance'
    else:
        responsive = item['clinical_significance']
    alteration = variant_name
    if alteration.lower().startswith(gene.lower()) is False:
        alteration = ' '.join([gene, alteration])
    drug_interaction_type = item['drug_interaction_type']
    if drug_interaction_type.lower() == 'combination':
        drug = drug.replace(',', '+')
    return {
        'responsive': responsive,
        'gene': gene,
        'vars': vars,
        'db_name': db_name,
        'drug': drug,
        'level': evidence_level,
        'cancer_type': cancer_type,
        'exon': exon,
        'status': status_en,
        'evidence_detail': [db_name, alteration, drug, evidence_type, match_degree, pmids]
    }


def reset_oncokb(item):
    db_name = 'oncokb'
    # 各数据库证据级别对应关系 https://mubu.com/doc/3LSWugo_cE
    # [数据库名, level, 肿瘤类型, 标准化肿瘤类型(汉语)]
    # level( "cgi", "Late trials" , "NSCLC", "肺癌")
    # 癌种是L列。当前肿瘤就是前面基本信息里获得的。
    # level(  "civic", "D: Preclinical evidence", "Ewing Sarcoma", "肺癌")
    # @霍  癌种是diseasename 列
    # level ( “oncokb”,  “R1”， “Non-Small Cell Lung Cancer”，  "肺癌")
    vars = []
    variant_name = item['Alteration']
    status_en, status_cn = reset_status(variant_name)
    exon = ''
    gene = item['Gene']
    drug = item['Drugs(s)']
    if status_cn == '未知':
        variant_names = variant_name.split(' ')
        if variant_name.startswith('Exon'):
            exon = 'E%s' % variant_names[1]
            status_en, status_cn = reset_status(variant_names[2])
        else:
            vars = [variant_name]
    Level, cancer_type = item['Level'], item['Cancer Type']  # 证据类型
    evidence_type = ', '.join([transfer_level(Level), cancer_type])
    # 敏感（responsive）的药物标蓝；耐药（resistance）
    if item['Level'].lower() == 'r1':
        responsive = 'resistance'
    else:
        responsive = 'responsive'
    alteration = variant_name
    if alteration.lower().startswith(gene.lower()) is False:
        alteration = ' '.join([gene, alteration])
    return {
        'responsive': responsive,
        'gene': gene,
        'vars': vars,
        'db_name': db_name,
        'drug': drug,
        'level': Level,
        'cancer_type': cancer_type,
        'exon': exon,
        'status': status_en,
        'evidence_detail': [db_name, alteration, drug, evidence_type, '完全匹配', item['PMIDs for drug']]
    }


def get_variant_knowledges():
    categories = []
    for i in range(variant_knowledge.nrows):
        j = 1
        cell_value = variant_knowledge.cell_value(i, j)
        if cell_value.startswith('rs') is False and i > 0:
            drug = '' if cell_value != '铂类药物' else '(顺铂、卡铂、奥沙利铂)'
            category = {'category': cell_value, 'drug': drug, 'genes': []}
            for k in range(i + 1, variant_knowledge.nrows):
                row_value = variant_knowledge.row_values(k)
                cell_value1 = row_value[j]
                if not cell_value1.startswith('rs'):
                    break
                gene = variant_knowledge.cell_value(k, j-1)
                if gene == '':
                    print variant_knowledge_name
                #     gene = variant_knowledge.cell_value(k-1, j-1)
                item = {
                    'gene': gene,
                    'rs': cell_value1,
                    'genotype': row_value[j+1],
                    'summary': row_value[j+2],
                    'introduction': row_value[j+3],
                    'level': row_value[j+4],
                    'category': category,
                }
                for x in rs_geno:
                    if x[0] == item['rs']:
                        if x[1] == item['genotype'] and item not in category['genes']:
                            category['genes'].append(item)
            if len(category['genes']) > 0:
                categories.append(category)
    return categories


def get_domain(gene, var_p, var_c):
    if CGI_mutation_analysis is not None:
        items = filter(lambda x: x['gene'] == gene and x['cdna'] == var_c, CGI_mutation_analysis)
        if len(items) != 1:
            items2 = filter(lambda x: x['gene'] == gene and x['protein'] == var_p, CGI_mutation_analysis)
            if len(items2) == 1:
                return items2[0]['Pfam_domain'].strip()
        else:
            return items[0]['Pfam_domain'].strip()
    return ''


def get_line(rId):
    img = get_img(img_info_path, rId)
    if img is not None:
        zoom = 17.7 / img['w']
        return int(zoom * img['h']) + 1
    return 0


def compare_level3(list1, list2):
    # 当相反证据量数量一致时，根据证据级别确定（证据级别的等级 1A＞1B＞2A＞2B＞3）
    eqs = ['1A', '1B', '2A', '2B', '3']
    for i in range(len(eqs)):
        eq = eqs[i]
        aa = filter(lambda x: x['level'] == eq, list1)
        for j in range(0, i+1):
            eq1 = eqs[j]
            bb = filter(lambda x: x['level'] == eq1, list2)
            if j < i and len(bb) > 0:
                return 2
            if len(aa) * len(bb) > 0:
                return 3
            if len(aa) > 0:
                return 1
            if len(bb) > 0:
                return 2
    return 3


def filter_rs_list(old_item):
    items = old_item['genes']
    rs_list = old_item['rs_list']
    for rs_item in rs_list:
        rs_list1 = []
        for item in items:
            if item['gene'] == rs_item['gene'] and item['rs'] == rs_item['rs'] and item['level'] == rs_item['level'] and item['category'] == rs_item['category']:
                rs_list1.append(item)
        rs_item['rs_list'] = rs_list1
    return rs_list


def concat_str(arr):
    if len(arr) == 1:
        return arr[0]
    tip1 = '、'.join(arr[:-1])
    if len(arr) > 1:
        return '合并'.join([tip1, arr[-1]])
    return tip1


def reset_sig(s_id):
    signature_cns = signature_cn.split('\n')
    for s_index, s in enumerate(signature_cns):
        if s.strip().upper() == s_id.upper():
            if s_index < len(signature_cns) - 1:
                text = signature_cns[s_index + 1].strip()
                if text == '原因未知':
                    return text
                return text.replace('推测由', '').replace('导致', '')
    return ''


def crop_img(input_url, output_url):
    if os.path.exists(input_url) is False:
        msg = 'file not exsits, %s' % input_url
        print msg
        return msg
    if os.path.exists(os.path.dirname(output_url)) is False:
        msg = 'folder not exists, %s' % os.path.dirname(output_url)
        print msg
        return msg
    img = Image.open(input_url)
    sp = img.size
    w, h = sp[0], sp[1]
    region = (0, h/3, w, h/3 * 2)
    if 'pie' in input_url:
        region = (w/3 - 60, 320, w / 3 * 2 + 200, h-320)
    if os.path.exists(output_url):
        os.remove(output_url)
    #裁切图片
    cropImg = img.crop(region)
    #保存裁切后的图片
    cropImg.save(output_url)
    return 'success'