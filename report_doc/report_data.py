#! /usr/bin/env python
# coding: utf-8
__author__ = 'huo'
import os
import time
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

from report_doc.tools.File import File
from report_doc.tools.Word import uniq_list, bm_index0
from report_doc.tools.get_real_level import FetchRealLevel

gray = 'E9E9E9'
gray_lighter = 'EEEEEE'
blue = '00ADEF'
white = 'FFFFFF'
red = 'ED1C24'
orange = 'F14623'
colors = ['2C3792', '3871C1', '50ADE5', '37AB9C', '27963C', '40B93C', '80CC28']

real_level = FetchRealLevel()
my_file = File()
immune_suggestion = my_file.read(r'immune_suggestion.txt', dict_name=r'data', sheet_name='immune_suggestion')
rs = my_file.read('3.rs.json', dict_name='data')
drugs = my_file.read('3.chem_durg_list.tsv', dict_name='data')  # category	drug	cancer
rs_genotype = my_file.read('3.rs.genotype.tsv', dict_name='data')
# variant_knowledge = my_file.read('3.chem_variant_knowledge.tsv', dict_name='data')
variant_knowledge = my_file.read(u'化疗多态位点证据列表V2 20180717.xlsx', dict_name='data', sheet_name='Sheet1')
imgs = my_file.read('img_info.json', dict_name='data')
hla = my_file.read('4.21.hla.tsv', dict_name='data')
variant_anno = my_file.read('0.2.2.variant_anno.maf.v2.xls', sheet_name='variant_anno.v2', dict_name='data')
gene_list12 = my_file.read('1.2gene_list.json', dict_name='data')
gene_list53 = my_file.read('5.3gene_list.xlsx', dict_name='data', sheet_name='Sheet2')
neoantigen = my_file.read('2.3.1_neoantigen.tsv', dict_name='data')
cnv_copynumber = my_file.read('cnv.copynumber.table.tsv', dict_name='data/cnv')
quantum_cellurity = my_file.read('0.2.3.quantum_cellurity.tsv', dict_name='data')
evidence_oncokb = my_file.read('allActionableVariants.txt', dict_name='data')
evidence_cgi = my_file.read('drug_prescription.tsv', dict_name='data')
evidence_civic = my_file.read('dependent/all_civic.180212.csv', dict_name='data')
signature_etiology = my_file.read('signature_etiology.csv', dict_name='data/signature')
gene_MoA = my_file.read('gene_MoA.tsv', dict_name='data')
recent_study = my_file.read('recent_study', dict_name='data')

evidence4 = []
for index in range(5):
    if os.path.exists('data/4.%d.evidence.txt' % index):
        evidence = my_file.read('4.%d.evidence.txt' % index, dict_name='data').split('\n')
    else:
        evidence = []
    evidence4.append(evidence)


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


def get_evidences(info, gene):
    db2 = info['evidence']
    if info['rId'].lower() == 'oncokb':
        return get_oncokb_evidence(db2, gene)
    evidences = []
    for i in range(len(db2)):
        # for j in range(len()):
        item = db2[i]
        evi = []
        if info['rId'].lower() == 'cgi':
            evi = get_cgi_evidence(item)
        elif info['rId'].lower() == 'civic':
            evi = get_civic_evidence(item)
        if len(evi) > 0:
            if evi[0] == gene:
                evidences.append(evi[1:])
    return evidences


def get_oncokb_evidence(info, gene):
    db2 = info.split('\n')
    line0 = []
    evidences = []
    for i in range(len(db2)):
        # for j in range(len()):
        line = db2[i].split('\t')
        if i == 0:
            line0 = line
        if len(line) == 10:
            gene_db1 = line[3]
            if gene_db1 == gene:
                info = {}
                for j in range(len(line)):
                    key = line0[j]
                    value = line[j]
                    info[key] = value
                alteration = line[4]  # 生物标志物
                Drugs = info['Drugs(s)']  # 药物
                Level, cancer_type = info['Level'], info['Cancer Type']  # 证据类型
                evidence_type = ', '.join([transfer_level(Level), cancer_type])
                match_degree = '完全匹配'  # 匹配程度
                pmids = info['PMIDs for drug']  # 证据来源
                evidences.append([alteration, Drugs, evidence_type, match_degree, pmids])
    return evidences


def get_cgi_evidence(info):
    # 其中Biomarker对应生物标志物，Drug对应药物，Tested_tumor+Evidence+Effect 对应证据类型，Source对应证据来源，匹配程度全部为完全匹配，后续有特殊需要标注的再行手动处理。
    alteration = info['BIOMARKER']  # 生物标志物
    gene_db1 = alteration.split(' ')[0]
    Drugs = info['DRUG']  # 药物
    evidence_type = ', '.join([info['TESTED_TUMOR'], info['EVIDENCE'], info['EFFECT']]) # 证据类型
    match_degree = '完全匹配'  # 匹配程度
    pmids = info['SOURCE']  # 证据来源
    return [gene_db1, alteration, Drugs, evidence_type, match_degree, pmids]


def get_civic_evidence(info):
    # 示例表格：
    # Variant_name对应生物标志物；
    # drug_name对应药物；
    # disease_name+evidence_level+clinical_significance+evidence_direction+rating=证据类型，
    #   其中clinical_significance为Resistance or Non-Response时，仅显示Resistance，
    #   evidence_direction 为surport时，不显示，为Do not surport是，用sensitivity（Do not surport）表示；
    # evidence_description+pubmed_id为证据描述，其中pubmed_id在evidence_description之后加括号表示（PMID）；
    # 匹配程度全部为完全匹配，后续有特殊需要标注的再行手动处理。
    alteration = info['variant_name']  # 生物标志物
    gene_db1 = info['genesymbol']
    Drugs = info['drug_names']  # 药物
    evidence_type = ', '.join([info['disease_name'], info['evidence_level'], info['clinical_significance'], info['evidence_direction'], info['rating']]) # 证据类型
    match_degree = '完全匹配'  # 匹配程度
    pmids = '%s(%s)' % (info['evidence_description'], info['pubmed_id'])  # 证据来源
    return [gene_db1, alteration, Drugs, evidence_type, match_degree, pmids]


def filter_db(gene=None, var=None, impact=None):
    db_table = variant_anno
    items = []
    keys = db_table.row_values(1)
    for i in range(1, db_table.nrows):
        gene_db = db_table.cell_value(i, 0)
        var_db = db_table.cell_value(i, 36)
        impact_db = db_table.cell_value(i, keys.index('IMPACT'))
        is_match = gene == gene_db
        if var is not None:
            is_match = is_match and var_db == var
        if impact is not None:
            is_match = impact == impact_db
        if is_match:
            info = {}
            for j in range(db_table.ncols):
                key = db_table.cell_value(1, j)
                value = db_table.cell_value(i, j)
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
        [u"（六）、最新研究进展解读说明", 2, 8, 23],
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
    title = '%s附录目录%s%s' % (title_cn, ' ' * 62,  title_en)
    title1 = '%s%s%s' % (title_cn, ' ' * (69+18),  title_en)
    titles = [title, title1]
    cats = get_catalog()
    for i in [0, 4, 9, 12, 19]:
        titles.append('%s%s%s' % (cats[i]['title'], ' ' * (69+20),  title_en))
    return titles


def get_immu(source):
    immune_table = immune_suggestion.split('\n')
    for im in immune_table:
        ims = im.split('\t')
        if ims[0] == source:
            return ims
    return [source, '', '']


# 选出部分重要抗原信息
def get_neoantigen():
    items1 = []
    for item in neoantigen:
        ref = item['RefAFF'].strip().strip('\n')
        try:
            ref = float(ref)
            if ref > 500:
                items1.append(item)
        except Exception, e:
            pass
    items1.sort(key=lambda x:x['MutRank'])
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
    item1 = {'gene': gene, 'fill': gray, 'text': 'EGFR无突变', 'copynumber': 0}
    items = filter_db(gene)
    for item in items:
        impact = item['IMPACT'].strip().upper()
        variant_type = item['Variant_Type']
        if impact == 'HIGH':
            if reset_status(variant_type) == 'gain':
                is_match, copynumber, status = get_copynumber(gene, 'gain')
                if is_match:
                    return {'gene': gene, 'fill': red, 'text': 'EGFRT突变合并扩增', 'copynumber': copynumber}
                return {'gene': gene, 'fill': gray, 'text': 'EGFR无扩增', 'copynumber': copynumber}
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
    items = []
    items2 = []
    for db_name in ['CGI', 'civic', 'Oncokb']:
        get_target_tip(items, items2, diagnose, db_name)
    print 'target_tip', len(items2)
    vars = []
    for item2 in items2:
        var = [item2, [], [], [], []]
        for item1 in items:
            if item1[0] == item2:
                for i in range(1, 5):
                    if item1[i] != '':
                        var[i].append(item1[i])
                var += item1[5:]
        vars.append(var)
    return vars


def get_target_tip(items, items2, diagnose, db_name):
    db_items = my_file.read('%s_evidence.csv' % db_name, dict_name='data/evidence')
    genes = []
    i = 0
    for db_item in db_items:
        # 各数据库证据级别对应关系 https://mubu.com/doc/3LSWugo_cE
        # [数据库名, level, 肿瘤类型, 标准化肿瘤类型(汉语)]
        # level( "cgi", "Late trials" , "NSCLC", "肺癌")
        # 癌种是L列。当前肿瘤就是前面基本信息里获得的。
        # level(  "civic", "D: Preclinical evidence", "Ewing Sarcoma", "肺癌")
        # @霍  癌种是diseasename 列
        # level ( “oncokb”,  “R1”， “Non-Small Cell Lung Cancer”，  "肺癌")
        if db_name == 'CGI':
            reset_item = reset_cgi(db_item, diagnose)
        elif db_name == 'civic':
            reset_item = reset_civic(db_item, diagnose)
        elif db_name == 'OncoKB':
            reset_item = reset_oncokb(db_item, diagnose)
        else:
            reset_item = None
        if reset_item is not None:
            gene = reset_item['gene']
            this_level = reset_item['level']
            drug = reset_item['drug']
            i += 1
            for v in reset_item['vars']:
                v = 'p.' + v
                if [gene, v] not in genes:
                    genes.append([gene, v])
                    for maf_item in filter_db(gene, v):
                        cellurity = get_quantum_cellurity(maf_item['Chromosome'], maf_item['Start_Position'])
                        col1 = '%s %s\n%s\n(%s肿瘤亚克隆)' % (gene, get_exon(maf_item), maf_item['HGVSp_Short'], cellurity)
                        item = get_drug(col1, gene, this_level, drug)
                        items.append(item)
                        if col1 not in items2:
                            items2.append(col1)

            if len(reset_item['vars']) == 0:
                items1 = filter_db(gene)
                is_match = False
                if len(items1) > 0:
                    info = items1[0]
                    exon = get_exon(info)
                    variant_type = info['Variant_Type']
                    col1 = ''
                    exon1 = reset_item['exon']
                    variant_type1 = reset_item['status']
                    if exon1 == exon and variant_type1[:3].upper() == variant_type:
                        is_match = True
                        col1 = '%s exon %s %s' % (gene, exon, reset_item['stauts'])
                    elif exon1 == '' and variant_type1 != '':
                        is_match, copynumber, status = get_copynumber(gene, variant_type1)
                        col1 = '%s%s(拷贝数%s)' % (gene, status, copynumber)
                    if is_match and len(col1) > 0:
                        item = get_drug(col1, gene, this_level, drug)
                        items.append(item)
                        if col1 not in items2:
                            items2.append(col1)
    return items, items2


def get_drug(col1, gene, this_level, drug):
    col6 = get_gene_MoA(gene)  # 原癌都是激活，抑癌都是失活。
    item = [col1, '', '', '', '',  col6, gene]
    if this_level == 'A':
        item[1] = drug
    elif this_level == 'B':
        item[2] = drug
    elif this_level == 'C':
        item[3] = drug
    elif this_level == 'D':
        item[4] = drug
    return item


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


def get_quantum_cellurity(chr, pos):
    for item in quantum_cellurity:
        if isinstance(chr, float):
            chr = str(int(chr))
        if isinstance(pos, float):
            pos = str(int(pos))
        # print item['Start'], pos, item['Chr'].lstrip('chr'), chr, type(pos), type(item['Start']), type(chr)
        if item['Chr'].lstrip('chr') == str(chr) and item['Start'] == pos:
            return '%.2f%%' % (float(item['Cellularity']) * 100)
    return ''


def get_data3(rs, drugs):
    items = rs['nodes']
    new_items = []
    for item in items:
        new_item = {}
        category = item['text']
        drug = filter(lambda x: x['category'] == category, drugs)
        new_item['category'] = category
        new_item['drug'] = drug[0]['drug']
        new_item['rs_list'] = get_variant_knowledge(category)
        new_items.append(new_item)
    return new_items


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
        if freq > 0.7 and item['impact'] == 'HIGH':
            return red
    return gray


def get_data45():
    items = [['Signature ID', '权重', '主要出现的癌种', '可能的病因', '评论']]
    tr1, tr2 = '', '该癌症可能由'
    for item in signature_etiology:
        s_id = item['Signature_ID']
        if s_id != 'unknown':
            tr1 += '特征%s所占权重为%s，推测由%s导致;' % (s_id, item['Weight'], item['Keyword'])
            tr2 += '%s、' % (item['Keyword'])
            items.append([s_id, item['Weight'], item['Cancer_types'], item['Proposed_aetiology'], item['Comments']])
    return uniq_list(items), tr1.rstrip(';'), tr2.rstrip('、') + '导致'


def get_gene_MoA(gene):
    # 原癌都是激活，抑癌都是失活
    # 共三类：原癌、抑癌、未知
    # Act : 原癌；Lof：抑癌， 其他： 未知
    for item in gene_MoA:
        if item['gene'] == gene:
            if item['gene_MoA'] == 'Act':
                return '原癌基因激活突变'
            if item['gene_MoA'] == 'Lof':
                return '抑癌基因失活变异'
    return '未知'


def reset_status(status):
    if status.lower() in ['amplification', 'gain', 'amp']:
        return 'gain', '扩增'
    if status.lower() in ['loss', 'deletions', 'deletion', 'del']:
        return 'loss', '缺失'
    return status, '未知'


def reset_cgi(e, diagnose):
    db_name = 'CGI'
    # 各数据库证据级别对应关系 https://mubu.com/doc/3LSWugo_cE
    # [数据库名, level, 肿瘤类型, 标准化肿瘤类型(汉语)]
    # level( "cgi", "Late trials" , "NSCLC", "肺癌")
    # 癌种是L列。当前肿瘤就是前面基本信息里获得的。
    # level(  "civic", "D: Preclinical evidence", "Ewing Sarcoma", "肺癌")
    # @霍  癌种是diseasename 列
    # level ( “oncokb”,  “R1”， “Non-Small Cell Lung Cancer”，  "肺癌")
    vars = []
    biomaker = e['BIOMARKER'].strip('"').split(' ')
    gene = biomaker[0]
    if len(biomaker) > 1:
        a = biomaker[1].split('(')
        if len(a) > 1:
            vars = a[1].rstrip(')').split(',')
    drug = e['DRUG']
    exon = ''
    status = ''
    this_level = real_level.level(db_name.lower(), e['EVIDENCE'], e['TESTED_TUMOR'], diagnose)
    if len(vars) == 0:
        if len(biomaker) == 4:
            exon = biomaker[2]
            status = biomaker[3]
        if len(biomaker) == 2:
            # D:\pythonproject\report\data\cnv\cnv.copynumber.table.tsv
            status = biomaker[1]
    return {
        'gene': gene,
        'vars': vars,
        'db_name': db_name,
        'drug': drug,
        'level': this_level,
        'exon': exon,
        'status': status
    }


def reset_civic(e, diagnose):
    db_name = 'civic'
    # 各数据库证据级别对应关系 https://mubu.com/doc/3LSWugo_cE
    # [数据库名, level, 肿瘤类型, 标准化肿瘤类型(汉语)]
    # level( "cgi", "Late trials" , "NSCLC", "肺癌")
    # 癌种是L列。当前肿瘤就是前面基本信息里获得的。
    # level(  "civic", "D: Preclinical evidence", "Ewing Sarcoma", "肺癌")
    # @霍  癌种是diseasename 列
    # level ( “oncokb”,  “R1”， “Non-Small Cell Lung Cancer”，  "肺癌")
    vars = []
    variant_name = e['variant_name']
    status_en, status_cn = reset_status(variant_name)
    exon = ''
    gene = e['genesymbol']
    drug = e['drug_names']
    this_level = real_level.level(db_name.lower(), e['evidence_level'], e['disease_name'], diagnose)
    if status_cn == '未知':
        variant_names = variant_name.split(' ')
        if variant_name.startswith('Ex'):
            vars = [gene, variant_names[-1]]
            exon = 'E%s' % variant_names[0].lstrip('Ex')
            status_en, status_cn = reset_status(variant_names[1])
        elif variant_name.startswith('Exon'):
            exon = 'E%s' % variant_names[1]
            status_en, status_cn = reset_status(variant_names[2])
        else:
            vars = [variant_name]

    return {
        'gene': gene,
        'vars': vars,
        'db_name': db_name,
        'drug': drug,
        'level': this_level,
        'exon': exon,
        'status': status_en
    }


def reset_oncokb(e, diagnose):
    db_name = 'oncokb'
    # 各数据库证据级别对应关系 https://mubu.com/doc/3LSWugo_cE
    # [数据库名, level, 肿瘤类型, 标准化肿瘤类型(汉语)]
    # level( "cgi", "Late trials" , "NSCLC", "肺癌")
    # 癌种是L列。当前肿瘤就是前面基本信息里获得的。
    # level(  "civic", "D: Preclinical evidence", "Ewing Sarcoma", "肺癌")
    # @霍  癌种是diseasename 列
    # level ( “oncokb”,  “R1”， “Non-Small Cell Lung Cancer”，  "肺癌")
    vars = []
    variant_name = e['Alteration']
    status_en, status_cn = reset_status(variant_name)
    exon = ''
    gene = e['Gene']
    drug = e['Drugs(s)']
    this_level = real_level.level(db_name.lower(), e['Level'], e['Cancer Type'], diagnose)
    if status_cn == '未知':
        variant_names = variant_name.split(' ')
        if variant_name.startswith('Exon'):
            exon = 'E%s' % variant_names[1]
            status_en, status_cn = reset_status(variant_names[2])
        else:
            vars = [variant_name]
    return {
        'gene': gene,
        'vars': vars,
        'db_name': db_name,
        'drug': drug,
        'level': this_level,
        'exon': exon,
        'status': status_en
    }


def get_rs_list(rs_list):
    items = []
    for rs in rs_list:
        item = {}
        item['rs'] = rs['text']
        item['genotypes'] = [r['text'] for r in rs['children']]
        items.append(item)
    return items


def get_variant_knowledge(category):
    # variant_knowledge gene	rs	genotype	introduction	summary
    items = []
    for i in range(variant_knowledge.nrows):
        j = 0
        cell_value = variant_knowledge.cell_value(i, j)
        if cell_value == category:
            print j, u'%s' % category
            for k in range(i + 1, variant_knowledge.nrows):
                row_value = variant_knowledge.row_values(k)
                cell_value1 = row_value[j]
                if not cell_value1.startswith('rs'):
                    break
                item = {
                    'gene': 'GENE%d%d' % (k, j),
                    'rs': cell_value1,
                    'genotype': row_value[j+1],
                    'summary': row_value[j+2],
                    'introduction': row_value[j+3],
                    'level': row_value[j+4],
                    'category': category
                }
                # print item['rs']
                items.append(item)
    return items