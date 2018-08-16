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
    img_info = get_img_info(base_dir, is_refresh=True)
    body = write_body(title_cn, title_en, data)
    titles = get_page_titles()
    pkgs1 = write_pkg_parts(img_info, body, title=titles)
    return pkgs1


def write_body(title_cn, title_en, data):
    chem_items, trs3 = get_data3(chem_durg_list)
    diagnose = data['patient_info']['diagnose']
    data['target_tips'] = get_target_tips(diagnose)
    data['chem_tip'] = ', '.join(trs3[1][1])
    body = ''
    body += write_cover(title_cn, title_en)
    body += write_catalog()
    body += write_chapter0(title_cn, data, 2)
    body += write_chapter1(data, 3)
    body += write_chapter2(4)
    body += write_chapter3(5, trs3, chem_items)
    body += write_chapter4(6)
    body += write_chapter5(7)
    body += write_backcover()
    return body


def write_cover(cn, en):
    cx = 3.76
    run_logo1 = r.picture(cx, rId='cover1', posOffset=[0, 0.5])
    run_logo2 = r.picture(cx, rId='cover1', posOffset=[7.8, 21.5])
    para = p.write(run_logo1)
    para += p.write(r.picture(21, rId='cover2', align=['center', 'center'], relativeFrom=['page', 'page']))
    para += p.write(run_logo2)
    para += p.write(p.set(sect_pr=set_page('A4')))
    return para


def write_catalog():
    # 目   录
    r.familyTheme = 'minorEastAsia'
    r.family_en = 'Times New Roman'
    r.family = ''
    para = p.write(p.set(spacing=[0, 0], jc='center', outline=3, line=24, rule='auto'), r.text("目录", size=18))
    catalogue = get_catalog()
    for cat in catalogue:
        # print index, cat[0], catalogue.index(cat)
        para = write_cat(cat, para, spacing=[0, 0.2], pos='8500')
    para += p.write(r.fldChar('end'))
    para += p.write(p.set(sect_pr=sect_pr_catalog))
    return para


def write_chapter0(title_cn, data, index):
    para = p.write(p.set(sect_pr=set_page('A4', footer='rIdFooter2', pgNumType_s=1, header='rIdHeader%d' % index)))
    para += p.write(p.set(spacing=[0, 0], jc='center', line=12, rule='auto'), r.text(title_cn, size=18, weight=1))
    para += write_patient_info(data['patient_info'])
    title = u'靶向治疗提示'
    para += p.h4(title)
    run = r.text('（')
    run += r.text('蓝色', color=blue)
    run += r.text('提示可能敏感药物，完整信息见附录）')
    para += p.write(p.set(jc='right'), run=run)
    target_tips_title = ['多组学发现', 'A级证据药物（CFDA、FDA、指南）', 'B级证据药物（专家共识）', 'C级证据药物（临床证据）', 'D级证据药物（临床前证据）']
    para += write_target_tip(target_tips_title, data['target_tips'])
    para += write_immun_tip()
    tips = [
        {'title': '化学治疗提示', 'text': '可能有效且毒副作用低的药物：%s' % data['chem_tip']},
        {'title': '最新研究进展治疗提示', 'sub_title': '（完整信息见附录）', 'text': recent_study},
        {'title': '检测方法', 'text': '检测技术：基于illumina Hiseq Xten平台，肿瘤组织外显子组80G、外周血外显子组10G\n相关局限性说明：由于肿瘤异质性等原因，本检测报告仅对本样本负责，患者诊疗决策需在临床医生指导下进行'},
    ]
    for tip in tips:
        para += p.h4(tip['title'])
        for t in tip['text'].split('\n'):
            if len(t) > 0:
                para += p.write(r.picture(cy=0.6, rId='tip0') + r.text(t, 9.5))
    para += p.write(p.set(sect_pr=set_page('A4', footer='rIdFooter2', pgNumType_s=1, header='rIdHeader%d' % index, type='continuous', cols=1)))
    return para


def write_chapter1(data, index):
    cats = get_catalog()[0: 4]
    para = ''
    title = ['多组学发现', 'A级证据药物（CFDA、FDA、指南）', 'B级证据药物（专家共识）', 'C级证据药物（临床证据）', 'D级证据药物（临床前证据）', '致癌模式']
    target_tips = data['target_tips']
    para += p.h4(cat=cats[0], spacing=[0, 0], color=white)
    para += p.h4(cat=cats[1], spacing=[0, 1]) + write_target_tip(title, data['target_tips'])
    para += p.h4(cat=cats[2])
    for i, item in enumerate(target_tips[1:]):
        para += p.write(p.set(shade=red, line=24), r.text('%d.变异：%s' % (i + 1, item['col1']), color=white, space=True))
        para += p.h4('（1）该基因变异所处保守结构域位置')
        para += p.write(p.set(jc='center', spacing=[3, 0]), run=r.picture(cx=15, rId='EGFR_struct', posOffset=[0, 0], align=['center', '']))
        para1, index1 = write_evidences(item)
        para += para1
        para += write_gene_info(item, index1)
        para += p.h4('（%d）基因突变保守结构域分布情况' % (index1 + 1))
        para += p.write(p.set(jc='center', spacing=[3, 0]), run=r.picture(cx=15, rId='EGFR_struct', posOffset=[0, 0], align=['center', '']))
        para += p.h4('（%d）基因突变各癌肿分布情况' % (index1 + 2))
        para += p.write(p.set(jc='center', spacing=[5, 0]), run=r.picture(12, rId='%s_distribution' % 'EGFR', posOffset=[0, 0], align=['center', '']))
        para += p.write() * 5
    para += page_br
    para += write_chapter13(cats[3], index)
    return para


def write_chapter2(index):
    n, start, bm0 = 4, 5, 453150350
    cats = get_catalog()[start-1: start + n]
    para = ''
    para += p.h4(cat=cats[0], spacing=[0, 0], color=white)
    chs = [
        {
            'title': p.h4(cat=cats[1], spacing=[0, 1]),
            'img_id': 'msi',
            'func': write_immun_tip1,
            'note': '注：INdel比例仅为MSI检测的17个参数之一',
            'infos': [
                {'title': '结果说明：', 'text': '该结果是通过肿瘤全外显子组数据，利用了一种基于TCGA（癌症基因图谱计划）数据集和引入多个突变特征开发的机器学习算法进行MSI（微卫星不稳定）状态评估方法得到的。该评估方法与传统方法相比，具有97.7%的准确性，是一种可靠的MSI检测方法。'},
                {'title': '检测意义：', 'text': 'FDA批准PD1抗体keytruda用于所有的MSI高不稳定的实体肿瘤。MSI高不稳定的肿瘤PD1治疗的客观缓解率在不同癌种和试验中稍有不用，大致在40-50%之间。MSI-H在结直肠癌、的分布。MSI是MMR（错配修复系统）的一个反2类型肿瘤细胞具有非常高的突变MSI-H/dMMR肿瘤更可能有效。'},
            ] },
        {
            'title': p.h4(cat=cats[2]),
            'img_id': 'tmb',
            'func': write_immun_tip2,
            'note': '注：NSCLC未经选择人群PD抗体有效率，具吸烟史为22%，无吸烟史为10%',
            'infos': [
                {'title': '结果说明：', 'text': '该结果是通过肿瘤外显子组检测得到的。全外显子组包含人体所有大约两万多个基因，大约有50M左右的区域。TMB肿瘤突变负荷指平均每M（兆）区域，肿瘤细胞发生的非同义突变的个数。同时，进一步将该结果与TCGA（肿瘤基因图谱）上相同癌种的TMB数据进行比对（如该癌种TCGA未收录，则与所有癌种的总集合进行比较），标注其TMB数值高于多少比例的人群。'},
                {'title': '检测意义：', 'text': 'TMB在多项临床研究中均被证明能够有效区分PD1抗体、CTLA4抗体等免疫检查位点抗体治疗是否有效的人群。综合型研究表明，在不同肿瘤中，不同患者的PD1抗体治疗有效性的差异55%可以由TMB的差异解释。TMB是不同肿瘤间体细胞突变量的评估。一般情况下，TMB越高，该肿瘤可能会拥有更多的肿瘤新生抗原，该肿瘤也越有可能在经过免疫检查位点抗体解除肿瘤免疫逃逸之后，被患者自身的免疫系统所识别，相关治疗在该患者身上也就越可能有效。'},
            ]
        }
    ]
    for ch in chs:
        para += write_chapter21(ch)
    para += p.h4(cat=cats[3])
    trs = [
        {'text': ['基因', '多肽', '亲和力（突变/正常）', '突变点（克隆组成）', 'HLA分子']},
    ]
    neoantigen = get_neoantigen()
    for n in neoantigen[:15]:
        trs.append({'text': [n['GeneSymbol'], n['MutPeptide'], '/'.join([n['MutAFF'], n['RefAFF']]), n['Mut'], n['HLA']]})
    # trs += [{'text': ['fff'] * 5}] *5
    borders=['top', 'bottom', 'left', 'right']
    pPr = p.set(spacing=[0.5, 0.5])
    para += table.write_jy1(trs, [1600, 1600, 2400, 2400, 1600], tc_borders=borders, th_borders=borders, th_color=gray, cell_color=gray_lighter, tc_pPr=pPr, th_pPr=pPr)
    para += p.write(r.text('注：仅显示部分重要新抗原信息', size=8.5))
    para += write_explain({'title': '结果说明：', 'text': '新抗原是指因肿瘤基因突变所导致的能够被该患者免疫系统HLA分子所识别，有潜力能够激活患者免疫系统的新生抗原，这是一组异常多或者特异性细胞治疗最核心的信息。新抗原预测一般通过外显子组测序得到该患者所有的息和HLA分子信息，预测各个突变位点与该患者HLA分子结合的行转录组检测，会明显降低新抗原预测的准确性。'})
    para += p.write(p.set(sect_pr=set_page('A4', header='rIdHeader%d' % index)))
    para += p.h4(cat=cats[4])
    para += p.h4('1、免疫治疗与肿瘤免疫')
    para_set = p.set(ind=['firstLine', 2])
    para += p.write(para_set, r.text('自2012年，约翰霍普金斯大学PD1抗体临床试验结果发表在《新英格兰医学杂志》之后，免疫治疗真正进入临床医生和产业界的视野，并迅速取代靶向治疗成为肿瘤治疗最新最主流的治疗研究方向。CTLA4、PD1抗体以惊人的产业化速度和疗效，迅速在多个癌种中获批，并在肺癌中成为一线治疗药物。PDL1表达、微卫星不稳定迅速成为有效协助患者药物筛选的分子标志物，肿瘤突变负荷TMB、新抗原负荷、肿瘤CD8+T淋巴细胞浸润状态等在各种数据中证明具有筛选有效患者的能力，然而，新型免疫治疗手段以超越基础研究的速度在发展，免疫治疗人群筛选和联合治疗时机选择远远没有到达完美，新的标志物层出不求'))
    para += p.write(para_set, r.text('肿瘤免疫治疗是通过协助免疫系统发挥被抑制或者缺失的免疫能力，进而实现肿瘤治疗的手段。更好的理解肿瘤免疫过程，能够协助我们对各种新型肿瘤免疫治疗手段和新型标志物形成更好的全局性理解，以便我们更好的采取与化疗、放疗、靶向治疗的联合治疗手段，优化免疫治疗时机和方式的选择。肿瘤免疫周期理论是由Mellman等人提出来的肿瘤免疫过程框架，目前已经成为肿瘤免疫学研究和临床应用的思维框架。'))
    para += p.h4('2、肿瘤免疫周期理论')
    para += p.write(r.picture(cy=10, rId='2.4.2', align=['center', ''], posOffset=[0, 0.5]))
    para += p.write(para_set) * 19
    para += p.write(para_set, r.text('肿瘤免疫治疗是通过协助免疫系统发挥被抑制或者缺失的免疫能力，进而实现肿瘤治疗的手段。更好的理解肿瘤免疫过程，能够协助我们对各种新型肿瘤免疫治疗手段和新型标志物形成更好的全局性理解，以便我们更好的采取与化疗、放疗、靶向治疗的联合治疗手段，优化免疫治疗时机和方式的选择。肿瘤免疫周期理论是由Mellman等人提出来的肿瘤免疫过程框架，目前已经成为肿瘤免疫学研究和临床应用的思维框架。'))
    para += p.h5('1）肿瘤细胞死亡并释放肿瘤特异抗原')
    para += p.write(para_set, r.text('从某种意义上说，肿瘤是由基因突变积累形成的疾病，同时，肿瘤基因突变也是肿瘤免疫过程发生的主要驱动因素。各种原因形成的肿瘤基因突变，部分情况下会让肿瘤细胞表达出与未突变基因不一样的蛋白质，其中部分突变蛋白的肽段会被身体的免疫系统识别为“外源”抗原，即所谓的“新抗原”。当肿瘤因为各种原因死亡后，含有新抗原的蛋白就会从肿瘤细胞中释放出来，进而驱动肿瘤免疫反应。所以，部分情况下，放疗、化疗、靶向治疗等治疗手段会跟免疫检查位点抗体治疗形成协同增强作用，也是因为该原因引起。最新部分研究甚至发现，不管PDL1表达状态，PD通路抗体联合化疗的疗效远高于单独化疗。微卫星高不稳定、肿瘤突变负荷高的肿瘤患者，一般情况下免疫检查位点抗体具有更好的疗效，也均是因为一般情况下，更高的突变负荷，意味着更多的新抗原可能性。'))
    para += p.h5('（2）抗原提呈细胞摄取并处理肿瘤新抗原')
    para += p.write(para_set, r.text('抗原提呈细胞摄取含有新抗原的突变蛋白，并将水解为可以结合到MHC（主要组织相容性复合物，又称为HLA人类白细胞抗原）的肽段。抗原提呈细胞需要有包括促炎症细胞因子和肿瘤细胞死亡释放的信号因子在内的多种免疫原性信号才能够有效启动，因此，不同免疫状态下的肿瘤，新抗原的提呈效率是不一样的。不同人的MHC具有异质性，所以，不同患者哪些新抗原能够被提呈是不一样的。目前通过生物信息手段，新抗原能够以一定的准确性被预测出来。目前有研究通过预测的新抗原进行个性化治疗疫苗设计，促进抗原提呈过程，并取得了相当惊人的效果。'))
    para += p.h5('（3）抗原提呈细胞进入淋巴结激活T细胞')
    para += p.write(para_set, r.text('（抗原提呈细胞被新抗原激活后，进入淋巴结激活初始T细胞。初始T细胞在共刺激分子的协同作用下诱导T细胞增殖并分化成为激活免疫杀伤的效应T细胞和抑制免疫杀伤的调节性T细胞，并通过两者比例的精确控制达成免疫反应性质确定和平衡。重组白细胞介素2（IL-2）即通过该环节调节肿瘤免疫反应。'))
    para += p.h5('（4）激活的T细胞离开淋巴结进入循环系统')
    para += p.write(para_set, r.text('激活的T细胞基于细胞黏附分子组分变化的原因，脱离淋巴结进入循环系统。'))
    para += p.h5('（5）T细胞穿过血管壁浸润到肿瘤微环境')
    para += p.write(para_set, r.text('肿瘤细胞死亡等原因引起的局部免疫反应会让周围组织跟炎症反应一样释放细胞因子和趋化因子，结合该部位血管内皮细胞由于炎症反应表达量增加的附着蛋白，让循环系统中的T细胞通过与血管内皮细胞锚定后进入肿瘤微环境中。初步研究结果表明，抗血管生成治疗与免疫治疗能够起到非常好的协同作用，部分临床试验中甚至提高将近一倍的治疗效果。'))
    para += p.h5('（6）T细胞通过特异性受体识别肿瘤细胞')
    para += p.write(para_set, r.text('肿瘤细胞和一般体细胞一样，在核糖体进行蛋白质翻译的过程中，会有一定比例的缺陷蛋白并会被水解成肽段，其中能够被MHC结合的潜在抗原会被提呈到细胞膜表面。T细胞通过特异性受体（TCR）识别肿瘤细胞表面的新抗原。有研究表明，部分肿瘤中，肿瘤细胞MHC缺失对于免疫逃逸的发生贡献重要意义。目前CAR-T和TCR-T是重点研究方向，并在部分肿瘤中发挥惊人的治疗效果。CAR-T治疗又称人工合成嵌合抗原受体T细胞治疗，直接取代T细胞识别肿瘤细胞的方式，而TCR-T则是通过筛选分离针对肿瘤特定抗原的TCR，将原来未能识别肿瘤细胞的T细胞改造成能够识别肿瘤细胞的T细胞。'))
    para += p.h5('（7）肿瘤细胞被T细胞识别并溶解消灭')
    para += p.write(para_set, r.text('当效应T细胞识别出肿瘤细胞表面的特异性抗原后，胞浆内储存的效应分子将朝目标方向释放，在不影响周围正常细胞的情况下溶解肿瘤细胞。CD8+T细胞是最重要的执行细胞杀伤的效应T细胞，所以，一般情况下，其肿瘤微环境浸润情况与免疫治疗疗效正相关。同时，PD1和CTLA-4是T细胞表面的抑制性受体，肿瘤细胞通过相应配体的表达，抑制T细胞的免疫杀伤作用。'))
    para += p.h4('3.肿瘤免疫表型理论')
    para += p.write(r.picture(cy=10, rId='2.4.3.1', align=['center', ''], posOffset=[0, 0.5]))
    para += p.write(para_set) * 19
    para += p.write(para_set, r.text('肿瘤免疫表型理论是由Mellman在肿瘤免疫周期的基础上，进一步根据最新的研究成果，细化发展出来的一套肿瘤免疫分型体系。该分型体系根据相应的生物学机制，将肿瘤分成免疫沙漠型（棕色）、免疫豁免型（蓝色）和炎症型（红色）三种，并进一步根据宿主基因、微生物组、环境因素、治疗药物和癌症共五个维度，将影响免疫原性的多种研究进展整合成如下图所示，与免疫治疗疗效和免疫原性相关的癌症-免疫设定点。'))
    para += p.write(r.picture(cy=8, rId='2.4.3.2', posOffset=[2, 0.5]))
    para += p.write(para_set) * 4
    para += p.write(p.set(ind=[24.5, 0]), r.text('癌症免疫设定点是指产生有效癌症免疫原性所需克服的阈值，为指导免疫治疗临床应用和研究提供一个系统性框架。该设定点可以理解为理解为刺激因子、抑制因子和TCR结合信号（T细胞抗原受体与新抗原、癌症相关抗原等癌症抗原的亲和力）的平衡。癌症免疫治疗主要是针对肿瘤部位，通过增加的刺激因子、减少抑制因子或者增加TCR结合信号这三种方式进行的。'))
    para += p.h4('4、免疫检查位点抗体疗联合传统治疗研究进展')
    para += p.write(para_set, r.text('免疫检查位点抗体单药治疗虽然在临床治疗中显示出广泛的抗癌效果，不同癌种的总体有效率基本在20%左右，即使通过多种标志物进行预测可一定程度上减少无效人群，但其获益人群一样有限。从肿瘤免疫过程上看，将免疫阻隔型和免疫沙漠型的“冷”肿瘤，转变成免疫炎症型的“热”肿瘤，扩大免疫检查位点抗体的获益人群，可以通过联合治疗的方式进行。对PDL1、MSI、TMB等各类型标志物预测检查位点抗体单药治疗效获益可能性较低的患者，最好进行检查位点抗体与其他治疗方式的联合治疗。多项研究发现放疗、化疗和靶向治疗等传统治疗联合免疫检查位点抗体治疗能够获得惊人的效果，中位生存期、无疾病进展生存期和有效率等疗效指标翻倍的情况。同时，值得注意的是，联合治疗会成倍甚至多倍的提高毒副反应。基于对肿瘤免疫过程理解的加深，特异性针对特定肿瘤免疫过程的免疫治疗是联合治疗的更优选择。'))
    para += p.h4('5、新型免疫治疗手段研究进展')
    para += p.write(para_set, r.text('免疫治疗是癌症治疗有史以来最激动人心的治疗领域，甚至有的医生认为，癌症免疫治疗让人类真正真正看到了癌症被治愈的希望。随着PD1抗体在各癌种中的攻城略地，新型的免疫治疗手段也展现出未来的王者之相。'))
    para += p.h5('（1）个性化癌症治疗疫苗')
    para += p.write(r.picture(cy=12, rId='2.4.5.1', align=['center', '']))
    para += p.write(para_set) * 21
    para += p.write(para_set, r.text('癌症疫苗，这种通过主动免疫去扩大肿瘤特异性T细胞反应的治疗方式，一直被认为是癌症免疫治疗的有效手段。尽管大家能够清晰看到癌症疫苗的合理性，但是，过去在临床方面的尝试都是不成功的。不同患者之间的肿瘤抗原具有强烈的多样性，因此，个性化癌症疫苗的发展是必要的。随着二代测序和生物信息工具的逐步完善，癌症疫苗的核心环节，新抗原预测逐渐成熟，该技术在最近的研究中取得突破性的进展，且由于安全性较好，是最值得跟进参与的新型癌症免疫治疗手段之一。'))
    para += p.h5('（2）免疫检查位点抑')
    para += p.write(para_set, r.text('肿瘤免疫检查位点不仅仅PD1和CTLA-4，还有至少几十种免疫检查位点。目前该领域，IDO抑制剂、LAG3抑制剂在早期临床试验中显示出相当好的疗效，与PD1联合用药的情况下，部分结果甚至成倍提升有效率，其中IDO抑制剂，已经进入三期临床临床试验（注：IDO抑制剂Epacadostat与Keytruda联用的关键三期临床试验ECHO-301失败）。'))
    para += p.h5('（3）CAR-T和TCR-T治疗')
    para += p.write(para_set, r.text('CAR-T和TCR-T都属于细胞治疗的范畴，主要通过对患者自身的T细胞进行工程化改造，让其能够发挥肿瘤细胞的杀灭功能。CAR-T，又称嵌合抗原受体T细胞治疗，是通过人工合成的受体使患者自身的T细胞能够进行肿瘤细胞识别，进而发挥肿瘤细胞杀伤效果。由于CAR-T细胞在实体瘤中的浸润能力相对较差，目前临床主要应用于血液肿瘤中。随着技术进展，如通过提升CAR-T中对增强T细胞浸润能力相关基因的表达，未来应该也能够在实体瘤治疗中发挥重要作用。TCR-T，又称T细胞识别受体（TCR）工程化改造T细胞治疗，是通过将对特定抗原亲和力强的TCR移植到患者自身的T细胞上使患者自身的T细胞发挥肿瘤细胞杀伤效果。其中特异性TCR-T，是指针对患者特异的新抗原进行设计的TCR-T治疗方式，是未来最有价值的癌症治疗手段，相比通过个性化疫苗诱导形成肿瘤杀伤T细胞，从原理上来说，特异性TCR-T治疗属于更靠后的免疫周期中的环节，可能具有更好的治疗效果。'))
    para += p.h5('（4）溶瘤病毒')
    para += p.write(para_set, r.text('溶瘤病毒是一群倾向于感染和杀伤肿瘤细胞的病毒。溶瘤病毒治疗是指将本身对身体伤害较低的溶瘤病毒经工程化改造减毒处理和治疗效果提升后，感染肿瘤患者的治疗方式。这种治疗思路和方法，是多年前发现和临床实践过的方法，且2005年中国CFDA批准了一种溶瘤腺病毒。但是，单药治疗效果有限，并未引起广泛关注。随着PD1抗体治疗的普及，临床研究发现，溶瘤病毒联合PD1治疗能够大幅度提高PD1抗体治疗的有效率，2015年，溶瘤病毒治疗T-Vec批准用于黑色素瘤。溶瘤病毒在提高PDL1表达、逆转肿瘤相关免疫抑制等多个层面，均能够与PD1抗体治疗形成非常好的协同效果。'))
    para += p.write(r.picture(cy=10, rId='2.4.5.2', align=['center', ''], posOffset=[0, 2.8]))
    para += p.write(p.set(sect_pr=set_page('A4', header='rIdHeader%d' % index)))
    return para


def write_chapter3(index, trs, chem_items):
    n, start, bm0 = 2, 10, 453150355
    cats = get_catalog()[start-1: start + n]
    c2 = ''
    for i in range(len(chem_items)):
        item = chem_items[i]
        cell = item['cell']
        row, col = cell[0], cell[1]
        c2 += p.h4('%d.%s（%s）' % (i+1, item['category'], item['drug']))
        data = [item['tr1'], '%s%s，%s' % (item['category'], trs[0][col], trs[row][0])]
        c2 += write_immun(data, jc='center', w=5000)
        c2 += p.write()
        text = 'PharmGKB'
        c2 += p.h4('（%d）                   %s药物基因组数据库（基因多态性相关证据）' % (0 + 1, text), runs=r.picture(cy=0.6, rId=text, posOffset=[1.3, 0.7]))
        c2 += p.write()
        genes = item['rs_list']
        c2 += write_gene_list3(item['genes'])
        c2 += p.write()
        for genes1 in genes:
            c2 += write_genotype(genes1['rs_list'], [1000, 8000])
    para = ''
    para += p.h4(cat=cats[0], spacing=[0, 0], color=white)
    para += p.h4(cat=cats[1], spacing=[0, 1])
    para += write_chemotherapy(trs, [800, 2400, 2400, 2400])
    para += p.write()
    para += write_explain({'title': '结果说明：', 'text': '该疗效预测汇总仅根据以下证据进行汇总。疗效预测证据主要来自CIVic数据库，并根据专家人工对相关证据进行梳理取舍。毒副作用证据及部分由基因多态性提供的疗效证据则来自于药物基因组数据库PharmGKB。采取该数据库二级以上的证据，并结合部分三级证据（不具备二级以上证据的情况下）。由于化学治疗的疗效影响因素较多，各个生物标志物的预测能力有限，且多个证据之间缺乏合理的证据平衡方式。目前初步采取多个证据之间目前采取均权投票的方式，即默认多个生物标志物的预测能力完全一致，仅根据各个证据的量进行评估,当相反证据量一致时，将证据级别纳入考量。'})
    para += p.h4(cat=cats[2])
    para += c2
    para += p.write(p.set(sect_pr=set_page('A4', header='rIdHeader%d' % index)))
    return para


def write_chapter4(index):
    cats = get_catalog()[12: 13 + 6]
    print cats[0]['title']
    para = ''
    para += p.h4(cat=cats[0], spacing=[0, 0], color=white)
    para += p.h4(cat=cats[1], spacing=[0, 1]) + write_chapter41()
    para += p.h4(cat=cats[2]) + write_chapter42()
    para += p.h4(cat=cats[3]) + write_chapter43()
    para += p.h4(cat=cats[4]) + write_chapter44()
    para += p.h4(cat=cats[5]) + write_chapter45()
    para += p.h4(cat=cats[6])
    para += p.write(r.text('癌症多组学临床检测是将多种维度的癌症全数据进行囊括的一种检测方式。随着癌症精准医疗和癌症真实世界研究高度发展，癌症循证医学证据以爆炸性的方式在扩大，整个组学甚至多个组学的整合性分析工具逐步出现。最新研究进展模块是皑医在持续的癌症最新研究进展文献调研跟进，多种组学整合型工具研究应用的基础上，在临床医生和相关生物医学研究人员共同决策的机制上，为了患者能够紧跟时代最前沿研究进展而设计的模块。该模块又进一步分成免疫治疗最新研究进展、肿瘤分子分型最新研究进展和其他最新研究进展这三类。由于该模块具有较强的前瞻性和不确定性，该模块的解读注释的使用务必在临床医生的主导下参考使用。'))
    para += p.write(p.set(sect_pr=set_page('A4', header='rIdHeader%d' % index)))
    return para


def write_chapter5(index):
    n = 4
    cats = get_catalog()[19: 20 + n]
    para = p.h4(cat=cats[0], spacing=[0, 0], color=white)
    para += p.h4(cat=cats[1], spacing=[0, 1]) + write_chapter51()
    para += p.h4(cat=cats[2])
    para += p.write(r.picture(12, rId='cnvanno_cicos', posOffset=[1.64, 1.49]))
    para += p.write(p.set(sect_pr=set_page('A4', header='rIdHeader%d' % index)))
    para += p.write(p.set(sect_pr=set_page()))
    para += p.h4(cat=cats[3]) + write_chapter53()
    para += p.h4(cat=cats[4])
    para += p.write(p.set(ind=['firstLine', 2]), r.text('本检测报告共涉及两个样本，肿瘤组织样本和外周血。本次检测使用安捷伦捕获外显子捕获探针Sureslect Human All Exon V6，目标测序数据量为肿瘤组织样本80G，外周血样本样本10G，实际数据量为肿瘤组织90.1G，外周血样本12.6G。经分析，肿瘤组织的预测肿瘤细胞纯度为81.93%，经初步分析共获得体细胞SNV 2257个，体细胞Indel303个，蛋白质编码区域突变838个，可能影响导致拷贝数变化的基因区域207个。'))
    para += p.write(p.set(jc='right', spacing=[3, 0]), r.text(format_time(frm="%Y-%m-%d")))
    para += p.write(p.set(sect_pr=set_page()))
    return para


def write_backcover():
    para = p.write(r.picture(rId='logo', posOffset=[0, 10], align=['center', ''], relativeFrom=['page', 'page']))
    para += p.write(p.set(spacing=[20, 13], jc='center'), r.text('重新定义癌症', 26, color=blue))
    para += p.write(r.text('皑医是由中国从业十年的第一批癌症精准医疗资深专家、医科院北京协和医学院的医学专家和中科院计算所的技术专家共同发起成立的一家癌症多组学数据临床解读公司。皑医致力于协助中国顶级肿瘤医生，以患者获益为中心，重新定义癌症，延长患者有质量生存时间。'))
    # para += p.write()

    imgs = [
        [
            {'id': 'logo', 'posOffset': 1},
            {'id': 'location', 'posOffset': 5.17},
        ],
        [
            {'id': 'website', 'posOffset': 1},
            {'id': 'phone', 'posOffset': 5}
        ]
    ]
    runs = [r.text('北京皑医科技有限公司北京市中关村科技园区大兴生物医药基地绿地启航3号楼1003', 9), r.text('https://aiyi.link 010-86399801', 9)]
    pos = [25.15, 26.32]
    for i in range(len(imgs)):
        infos = imgs[i]
        posy = pos[i]
        run = runs[i]
        for info in infos:
            run += r.picture(cy=0.6, rId=info['id'], posOffset=[info['posOffset']+0.5, posy], relativeFrom=['page', 'page'])
        para += p.write(p.set(spacing=[1, 0.5]), run)
    para += set_page(header='rIdHeader8', footer='rIdFooter1')
    return para


def write_explains(content):
    para = ''
    for text in content.split('\n'):
        texts = text.split('：')
        if len(texts) > 1:
            para += write_explain({'title': '%s：' % texts[0], 'text': texts[1]})
        else:
            para += p.write(r.text(text))
    return para


def write_chapter13(cat, index):
    para = p.h4(cat=cat)
    para += p.h4('1.靶向治疗与驱动基因')
    para += p.write(p.set(ind=['firstLine', 2]), r.text('靶向治疗药物是针对特定的肿瘤发生发展相关特定基因设计的药物。传统化疗主要针对快速分裂的细胞，既杀伤肿瘤细胞、又杀伤正常细胞，毒副作用大，而靶向治疗更精准的针对肿瘤特定特征或者肿瘤微环境，所以毒副作用相对较低。'))
    para += p.write(p.set(ind=['firstLine', 2]), r.text('实际临床应用中，靶向药初步可以分成针对肿瘤抗血管生成及多靶点相关的靶向药和针对肿瘤细胞特定基因变异的靶向药。抗血管生成及多靶点相关的靶向药现阶段大多没有特定的基因变异可以预测其疗效。针对肿瘤细胞特定基因变异的靶向药，一般情况下仅对该类突变患患者有效。此类特定基因，包含在肿瘤驱动基因范畴中。'))
    para += p.h4('2.驱动基因及突变形式说明')
    para += p.h4('（1）驱动基因说明')
    para += p.write(p.set(ind=['firstLine', 2]), r.text('在肿瘤发生发展中扮演重要角色，能够“驱动”癌症疾病进程的基因称为驱动基因。乘客基因则是指对肿瘤发生发展重要性不高的基因，但是，重要性目前只是一个相对概念而不是绝对概念，所以，虽目前有多种方式进行驱动基因的鉴定，但是并没有统一的完整标准。驱动基因又分成发生激活突变后具有促进癌症发生发展的原癌基因和功能正常情况下抑制癌症发生的抑癌基因。驱动基因上的基因变异，可以是驱动突变，也可以是乘客突变。驱动突变可以是原癌基因的激活突变，也可以是抑癌基因的失活突变。一般情况下，肿瘤靶向治疗针对驱动基因中的原癌基因激活突变进行抑制，如EGFR Tkis抑制EGFR突变，或者针对抑癌基因的失活突变进行相关信号通路的协同致死，如PARP抑制剂治疗BRCA1、2基因变异肿瘤。'))
    para += p.h4('（2）原癌基因和激活突变')
    para += p.write(p.set(ind=['firstLine', 2]), r.text('原癌基因指正常功能情况下，在细胞信号传导等多个层面扮演重要角色，但是发生激活突变后会促进癌症发生发展的基因。激活突变一般情况下发生在原癌基因经常发生突变的热点位置上，如下图的PIK3CA和IDH1基因，且突变以错义突变为主。'))
    para += p.h4('（3）抑癌基因和失活突变')
    para += p.write(p.set(ind=['firstLine', 2]), r.text('抑癌基因指正常功能下，扮演着DNA修复等抑制癌症发生发展过程的基因。抑癌基因发生失活突变会导致身体抑制癌症的功能降低。抑癌突变一般情况下热点突变较少，可发生在基因近乎任何区域，如下图的RB1和VHL基因，且会出现更多的截断突变。'))
    para += p.write(p.set(jc='center'), run=r.picture(13.34, rId='1.3.3'))
    para += p.write() * 7
    para += p.h4('3.肿瘤数据解读证据级别说明')
    para += p.write(p.set(ind=['firstLine', 2]), r.text('通过二代测序技术，特别是本报告采用的组学检测技术，每个肿瘤患者会找到几十个、几百个甚至于几千个肿瘤基因变异。不同基因变异具有不同的临床指导意义。美国临床肿瘤协会（ASCO）、美国病理学家联合学会会（CAP）和分子病理协会（AMP）共同发布了相关的标准和指南。该指南首先把变异根据临床意义等级分成四类：等级Ⅰ，强临床意义的变异；等级Ⅱ，潜在临床意义的变异；等级Ⅲ，不清楚临床意义的变异；等级Ⅳ，良性和可能良性的变异。其中，又根据变异的证据级别，等级Ⅰ和等级Ⅱ的变异进一步细化分成Level A、B、C、D四个级别。'))
    para += p.write(p.set(ind=['firstLine', 4]), r.text('Level A：1、针对某一特定癌症，经过FDA批准的；2，针对某一特定癌种，专业指南推荐的；'))
    para += p.write(p.set(ind=['firstLine', 4]), r.text('Level B：基于高水平研究，相关领域专家意见一致；'))
    para += p.write(p.set(ind=['firstLine', 4]), r.text('Level C：1、同一分子标志物，FDA批准用于其他癌症；2、作为临床试验纳入标准'))
    para += p.write(p.set(ind=['firstLine', 4]), r.text('Level D：临床前研究，结果不确定'))
    para += write_db_info()
    para += p.write(p.set(sect_pr=set_page(header='rIdHeader%d' % index)))
    return para


def write_chapter21(ch):
    para = ch['title']
    para += ch['func']('center')
    para += p.write(p.set(jc='center', spacing=[5.5, 0]), run=r.picture(cy=7.3 * 0.5, rId=ch['img_id'], posOffset=[0, 0], align=['center', '']))
    para += p.write(p.set(jc='center'), r.text(ch['note'], size=8.5))
    for i in ch['infos']:
        para += write_explain(i)
    return para


def write_chapter41():
    para41, reds, oranges = write_41()
    tr1 = '无致病位点， DDR信号通路未激活'
    tr2 = 'PD1等免疫检查位点抗体可能无效'
    if reds + oranges> 0:
        tr1 = ''
        if reds > 0:
            tr1 += 'DDR基因明确致病位点突变%d个' % reds
        if oranges > 0:
            if len(tr1) >0:
                tr1 += '，'
            tr1 += '预测致病位点突变%d个' % oranges
        tr1 += '，DDR信号通路激活。'
        tr2 = 'PD1等免疫检查位点抗体可能有效'
    para = write_immun([tr1, tr2], jc='center', h0=1600, w=3600)
    para += p.write()
    para += para41
    para += p.write()
    run = r.text('红色，', color=red, size=9)
    run += r.text('表示明确致病突变位点；', size=9)
    run += r.text('橙红色，', color=orange, size=9)
    run += r.text('表示预测致病突变位点；', size=9)
    run += r.text('未发生突变用灰色表示', size=9)
    para += p.write(p.set(spacing=[0.5, 0.5]), run)
    para += write_explain({'title': '结果说明：', 'text': 'DDR基因突变患者更可能从免疫检查位点抗体治疗中获益。细胞内正常的代谢活动与环境因素均会引起DNA损伤，初步估算，每个正常细胞每天产生1000-1000000处分子损伤。肿瘤细胞由于基因组等异常，大部分情况下DNA损伤的速率远高于正常细胞。DNA损伤反应基因主要包括错配修复、同源重组修复等DNA修复相关基因（DDR基因），也包括细胞周期检查点、染色质重塑等其他基因，据一篇专家手工注释研究表明，DDR基因大约有450个。DNA损伤反应跟免疫系统之间具有非同寻常的关系，如召集免疫细胞聚集，对T细胞杀伤更敏感等。以上是根据目前最新研究，将已有证据证明相关基因突变与免疫检查位点疗效直接相关的DDR基因子集。该基因列表可能会随着DDR基因的研究进展逐步扩大。'})
    para += write_evidence4(1)
    return para


def write_chapter42():
    # data = ['DDR基因明确致病位点突变1个，预测致病位点突变4个，DDR信号通路激活。', 'PD1等免疫检查位点抗体可能有效']
    # 1,  如果HLA-B66，或者HLA-B*15：01，疗效较差
    # 2， 如果LA-B44，较好，  如果三个均杂合（A1！=A2， B1！=B2， C1！=C2）， 疗效较好。
    # 3，其它，疗效中等。
    para = write_hla(hla)
    para += p.write()
    para += write_explain({'title': '结果说明：', 'text': 'HLA分型与免疫治疗疗效高度相关。HLA(human lymphocyte antigen ，人类淋巴细胞抗原)，是编码人类的主要组织相容性复合体（MHC）的基因。HLA是免疫系统区分自身和异体物质的基础。HLA主要包括HLA Ⅰ类分子和Ⅱ分子。HLAⅠ类分子又进一步细化分成A、B、C三个基因。特定的超型，如HLA-B44，与免疫检查点抗体治疗疗效好相关；HLA-B66（包括HLA-B*15：01），与免疫检查点抗体治疗疗效差相关。HLA Ⅰ类三个基因均杂合，免疫检查点抗体治疗反应更好。HLA杂合缺失的基因相关的新抗原可能在个性化治疗疫苗或者特异性细胞治疗中无效。'})
    para += p.write(p.set(spacing=[0.5, 1]), r.text('相关循证医学证据：'))
    para += p.write(p.set(spacing=[0, 0.2], shade=gray), r.text('1535例晚期癌症患者接受免疫检查点抗体治疗的HLA Ⅰ类分子分型整合分析'))
    para += p.write(p.set(spacing=[0, 1], shade=gray), r.text('（A）HLAⅠ类分子三个基因位置至少一个杂合的患者，总生存较短，风险比较高（队列1：杂合282名，至少一个纯合87名，HR=1.4（1.02-1.9），P=0.036；队列2：杂合899名，至少一个纯合267名，HR=1.31（1.03-1.7），P=0.028）；'))
    para += p.write(p.set(spacing=[0, 1], shade=gray), r.text('（B）黑色素瘤患者中，B44超型的患者，总生存较长，风险比较低（队列1：HLA-B44（+）55名；HLA-B44（-）109名，HR=0.5（0.32-0.76），P=0.001；队列2：HLA-B44（+）44名；HLA-B44（-）106名，HR=0.32（0.09-1.1），P=0.05）；'))
    para += p.write(p.set(spacing=[0, 1], shade=gray), r.text('（C）黑色素瘤患者中，B62超型（包括HLA-B*15：01）的患者，总生存较短，风险比较高。整合分析中，两个队列324名黑色素瘤患者中，B62超型患者风险比较高（HR=2.29 (1.40–3.74) P=0.0007）；类似于B62超型的B*15:01风险比较高（HLA-B*15:01（+）142名，HLA-B*15:01（-）22名，HR=2.21 (1.33–3.70)，P=0.002）（Chowell D,et al. [J]. Science, 2018, 359(6375): 582-587.）'))
    return para


def write_chapter43():
    # MDM2  查CNV表，看是否为GAIN
    # MDM4同MDM2
    # DNMT3A查maf表，看是否有突变
    # EGFR是查maf表和CNV表。
    # 11q13标红：仅当CCND1，FGF3，FGF4，FGF19全部扩增（GAIN）
    # 扩增=GAIN=amplication
    items = get_data43()
    tr1, tr2 = '', 'PD1等免疫检查位点抗体可能无超进展风险'
    for item in items:
        if item['fill'] == red:
            tr1 += '%s(拷贝数%s),' % (item['text'], item['copynumber'])
            tr2 = 'PD1等免疫检查位点抗体可能有超进展风险'
    trs = [tr1.rstrip(','), tr2]
    para = write_immun(trs, jc='center')
    para += p.write()
    para += write_43(items)
    para += p.write()
    para += write_explain({'title': '结果说明：', 'text': 'MDM2、MDM4基因扩增、EGFR变异（突变或扩增）、DNMT3A突变和11q13(CCND1、FGF3、FGF4、FGF19)扩增与PD1抗体治疗超进展，即治疗之后，肿瘤没有缩小，反而加速增大。患者出现以上基因异常事件时，需要相对慎重选择PD1抗体治疗。'})
    para += write_evidence4(3)
    return para


def write_chapter44():
    # 纯合失活突变：
    # 纯合：突变频率 >0.7  ，突变频率=AS/AQ
    # 失活：跟恶性突变是一样的，是指maf中的impact为HIGH
    genes = ['B2M', 'JAK1', 'JAK2', 'PTEN']
    items = []
    genes_red = []
    tr1 = '未发现免疫治疗耐药相关基因突变'
    tr2 = 'PD1等免疫检查位点抗体可能有效'
    for gene in genes:
        fill = get_data44(gene)
        text = '未'
        if fill == red:
            text = ''
            genes_red.append(gene)
        items.append({'fill': fill, 'text': '%s%s发生纯合失活突变' % (gene, text)})
    if len(genes_red) >0:
        tr1 = '发现%s基因发生纯合失活突变' % ('、'.join(genes_red))
        tr2 = 'PD1等免疫检查位点抗体可能无效'
    data = [tr1, tr2]
    para = write_immun(data, jc='center', h0=1600, w=3600)
    para += p.write()
    para += write_43(items, ws=[9600/4] * 4)
    para += p.write()
    para += write_explain({'title': '结果说明：', 'text': 'B2M等基因发生特定类型基因变异，PD1抗体治疗等免疫治疗可能会发生耐药。免疫治疗耐药可以由多种因素引起，以上基因通过不同机制导致免疫治疗耐药。B2M基因纯合失活突变，主要通过损害抗原提呈机制使免疫治疗耐药。JAK1、JAK2基因的纯合失活突变，则是通过损害效应T细胞杀伤肿瘤细胞的信号通路（γ干扰素通路）导致免疫治疗耐药。PTEN基因表达缺失或者纯合失活突变，则可能是通过影响T细胞浸润使免疫治疗耐药。'})
    para += write_evidence4(4)
    return para


def write_chapter45():
    items, tr1, tr2 = get_data45()
    data = [tr1, tr2]
    para = write_immun(data, jc='center', w=5000)
    para += p.write()
    para += p.write(p.set(spacing=[2, 0]), r.picture(16, rId='4.5.1signature', posOffset=[0, 0.1], align=['center', '']))
    para += p.write() * 9
    para += write_evidence(items[1:], titles=items[0])
    para += p.write(r.br('page'))
    para += p.write(p.set(spacing=[2, 0]), r.picture(7.5, rId='4.5.2signature_pie', posOffset=[0, 0.6]))
    para += write_explain({"title": '结果说明：', 'text': '体细胞突变存在于人体的所有细胞中并且贯穿整个生命。它们是多种突变过程的结果，包括DNA复制机制内在的轻微错误，外源或内源诱变剂暴露，DNA酶促修饰和DNA修复缺陷。不同的突变过程产生独特的突变类型组合，称为“突变特征”。在过去的几年中，大规模的分析研究揭示了许多人类癌症类型的突变特征。目前这组突变特征是基于对40种不同类型的人类癌症中的10,952个外显子组和1,048个全基因组的分析得到的。使用六个取代亚型显示每个标记的概况：C> A，C> G，C> T，T> A，T> C和T> G，进而向前先后各延伸一个碱基，每个碱基有4种可能，所以，总共就有96个三核苷酸的突变类型。现在已经明确的，总共有30种明确注释的“突变特征”，每种特征都有对应的发生机制，如吸烟、错配修复缺陷、同源重组修复缺陷等。'}, ind=[23, 0])
    return para


def write_chapter51():
    para = ''
    ws = [1800, 1800, 1600, 1600, 1800]
    pPr = p.set(jc='left', spacing=[0.5, 0.5])
    titles = ['基因突变（肿瘤细胞克隆组成）', '变异类型', '所在结构域', '功能预测', '驱动基因类型及致癌性']
    data = filter_db(impact='HIGH')
    trs = write_thead51(titles, pPr=pPr, ws=ws)
    for k in range(len(data)):
        tcs = ''
        size = 9
        item1 = data[k]
        gene = item1['Hugo_Symbol']
        variant = item1['HGVSp_Short']
        domain = get_domain(gene, variant, item1['HGVSc'])
        cellurity = get_quantum_cellurity(item1['Chromosome'], item1['Start_Position'])
        col1 = '%s %s %s' % (gene, variant, cellurity)
        col5 = get_gene_MoA(gene)
        item = [col1, item1['Variant_Classification'], domain, '有害突变', col5]
        for i in range(len(item)):
            t = item[i]
            run = r.text(t, size=size, weight=0)
            tcs += tc.write(p.write(pPr, run), tc.set(w=ws[i], tcBorders=['top', 'bottom', 'left', 'right'], color=gray_lighter))
        trs += tr.write(tcs)
    para += table.write(trs, ws=ws, bdColor=gray_lighter)
    para += p.write()
    return para


def write_chapter53():
    para = p.write(r.text('评估肿瘤重点信号通路的状态是从全局对癌症进行理解的重要方式。肿瘤相关驱动基因一般都处在肿瘤的重点信号通路上。该样本相关的肿瘤驱动基因变异和非驱动基因变异究竟累及哪些信号通路，信号通路之间各个基因的关系如何，针对相关信号通路的治疗方式以及各个信号通路上下游的信号通路的分布，都能够协助医生和患者理解肿瘤所处状态。'))
    para += p.write()
    chapters = [
        {
            'title': 'MAPK信号通路变异情况',
            'img_id': 'hsa04010.MAPK_signaling_pathway',
        },
        {
            'title': 'PI3K/AKT/mTOR信号通路变异情况',
            'img_id': 'hsa04150.mTOR_signaling_pathway',
        },
        {
            'title': '细胞凋亡信号通路变异情况',
            'img_id': 'hsa04210.Apoptosis',
        },

        {
            'title': 'WNT通路变异情况',
            'img_id': 'hsa04310.Wnt_signaling_pathway'
        },
        {
            'title': 'Hedgeho通路变异情况',
            'img_id': 'hsa04340.Hedgehog_signaling_pathway'
        },
        {
            'title': 'JAK/STAT通路变异情况',
            'img_id': 'hsa04630.Jak-STAT_signaling_pathway',
        },
    ]
    for i in range(gene_list53.nrows):
        title = gene_list53.cell_value(i, 0)
        ch = {'title': '%d.%s' % (i + 1, title), 'para': gene_list53.cell_value(i, 1)}
        genes = []
        for ch1 in chapters:
            if title in ch1['title']:
                ch['img_id'] = ch1['img_id']
        for j in range(3, gene_list53.ncols):
            v = gene_list53.cell_value(i, j)
            if v.strip() != '':
                genes.append(v)
        ch['genes'] = genes
        para += write_chapter5311(ch, i)
    return para


# Part5 . 3 通路涉及重点基因变异情况
def write_chapter5311(ch, index):
    para = p.write(r.text(ch['title']))
    para += con1
    para += p.write()
    para += p.write(r.picture(cx=7.8, rId='5.3.%d.1' % (index + 1)) + r.br())
    title = '' if 'img_id' not in ch else '（1）'
    para += p.write(r.text('%s通路涉及重点基因变异情况：' % title))
    para += write_genes((ch['genes']), 5, 5000, 'right')
    row = int(math.ceil(float(len(ch['genes'])) / 5))
    if row > 5:
    #     # para += p.write() * 3
        para += con2 + con1
    para += write_explains(ch['para'])
    if row > 5:
        para += con1
    else:
        para += con2
    if 'img_id' in ch:
        para += con1
        para += p.write(r.text('（2）KEGG完整信号通路基因突变注释图'))
        para += p.write(r.picture(rId=ch['img_id'], zoom=0.22, posOffset=[0, 0.5]))
        para += con1
        if index < 10:
            para += p.write(p.set(sect_pr=set_page()))
    return para


def write_hla(data, w=4000):
    if len(data) == 0:
        return ''
    # 1,  如果HLA-B66，或者HLA-B*15：01，疗效较差
    # 2， 如果LA-B44，较好，  如果三个均杂合（A1！=A2， B1！=B2， C1！=C2）， 疗效较好。
    # 3，其它，疗效中等。
    item = data[0]
    if item['B1'].startswith('B*66') or item['B2'].startswith('B*66') or 'B*15:01' in [item['B1'], item['B2']]:
        a_word = '疗效较差'
    elif item['B1'].startswith('B*44') or item['B2'].startswith('B*44'):
        a_word = '疗效较好'
    else:
        a_word = '疗效中等'
    data.append('PD1等免疫检查位点抗体可能%s' % a_word)
    texts = []
    for j in ['A', 'B', 'C']:
        h1 = item['%s1' % j].replace('*', '')
        h2 = item['%s2' % j].replace('*', '')
        texts.append('HLA-%s HLA-%s' % (h1, h2))
    trs2 = write_tr1('\n'.join(texts))
    trs2 += write_tr2('PD1等免疫检查位点抗体可能%s' % a_word)
    table_str = table.write(trs2, ws=[w], jc='center', bdColor=blue)
    return table_str


def write_evidences(item):
    infos = [
        {'rId': 'Oncokb', 'cx': 3, 'cy': 1, 'posy': 1, 'evidence': evidence_oncokb},
        {'rId': 'CGI', 'cx': 5.74, 'cy': 1.11, 'evidence': evidence_cgi},
        {'rId': 'CIVic', 'cx': 4.23, 'cy': 1.8, 'evidence': evidence_civic},
        {'rId': 'logo', 'text': '其他', 'cx': 1.35, 'cy': 1.28},
    ]
    para = ''
    index = 2
    for i in range(len(infos)):
        info = infos[i]
        rId = info['rId']
        if rId.upper() in item:
            evidences = item[rId.upper()]
            if len(evidences) > 0:
                text = rId if 'text' not in info else info['text']
                run = r.text('（%d） ' % (i + 1), 12, space=True, weight=True)
                run += r.picture(cy=0.5, rId=info['rId'], text_wrapping='inline')
                run += r.text(' %s数据库证据' % (text), 12, space=True, weight=True)
                para += p.write(p.set(spacing=[1.5, 1.5]), run)
                para += write_evidence(evidences)
                index += 1
    return para, index


def write_db_info():
    para = p.h4('4.肿瘤解读数据库相关说明')
    infos = [
        {'rId': 'CIVic', 'cx': 4.23, 'cy': 1.8, 'off_y': 0,
         'para': [
             {'text': 'CIVic是一个基于开放用户提供内容，领域专家进行审核的精准医疗知识社区。领域专家在社区中起着至关重要的作用。领域专家通常是科学家或医生，具有高级学位（PHD或MD级），并展示了与癌症精确医学知识相关的专长（发表记录）。该社区的知识库将变异分成五类：A级，已验证，已批准或者已形成共识；B级，临床证据，来源于临床试验证据或者其他患者来源证据支持；C级：案例研究，来源于临床的案例报道；D级：临床前证据，体内或体外生物学支持证据；E级：非直接证据。'}
         ]},
        {'rId': 'Oncokb', 'cx': 3, 'cy': 1,
         'para': [
             {'text': ' OncoKB是一个肿瘤精准医学数据库，包含特定的癌症基因变异效应和治疗提示的信息。这是由Memorial Sloan Kettering癌症中心（MSK）的MarieJosée和Henry R. Kravis分子肿瘤学中心的知识系统小组与Quest Diagnostics合作开发和维护的。由MSK的临床人员、研究人员和教学人员共同组成的协作网络进行专家手工注释，OncoKB包含有关477种癌症基因特定变异的详细信息。这些信息具有各种来源，如FDA，NCCN或ASCO等指南，和ClinicalTrials.gov和科学文献等。对于每一种变异，该数据库均会专家手工注释生物学效应、患病率、预后信息和治疗影响。治疗信息使用自行开发的证据级别系统进行分类。该证据级别共分为四类：FDA批准、专家共识、强有力临床证据和强有力生物学证据四种，其分类与ASCO、AMP和CAP共同发布的指南一一对应。'}
         ]},
        {'rId': 'CGI', 'cx': 5.74, 'cy': 1.11,
         # 'para': 'CGI(癌症基因组解释器)被设计用于识别驱动肿瘤发展，可能可治疗的变异。',
         'para': [
             {
                 'text': 'CGI(癌症基因组解释器)被设计用于识别驱动肿瘤发展，可能可治疗的变异。CGI依赖于多个数据源和计算机算法，依照不同的证据级别对变异进行分类。这是一个由巴塞罗那生物医学基因组实验室维护开发的一个癌症基因组注释工具。该注释工具相关的数据库大致将变异分为以下五类：临床指南、晚期临床试验、早期临床试验、临床案例报道和临床前数据这五类。由于人类知识积累有限，大部分基因变异是否为致癌变异均是未知的。该注释工具的核心特色是一个针对未知是否致癌的变异进行致癌性预测的工具，oncodriverMUT。该工具结合了大规模肿瘤基因组数据（来自28个癌种的6792个样本）和大规模正常人基因组数据（60706个未经选择的正常人样本）中突变的分布，并结合突变功能预测的多种规则，对未知变异的致癌性进行预测。未知突变经预测，分类为预测为驱动突变级别1、预测为驱动突变级别2和预测为乘客突变。',
                 'pic': {'rId': '1.3.4.1', 'cx': 14.1, 'cy': 8.57, 'off_y': 0}
             },
             {'text': ''},
             {'text': ''},
             {'text': ''},
             {'text': ''},
             {'text': ''},
             {'text': ''},
             {
                 'text': '通过对1077已知致癌变异、2819个癌症易感变异、241个癌症基因上已知为良性的PAM（影响蛋白功能突变）和1006个癌症基因上经常出现在一般人群中的PAM的验证，发现该算法的准确性为0.91。',
                 'pic': {'rId': '1.3.4.2', 'cx': 13.76, 'cy': 7.54}
             },
         ]},
    ]
    for i in range(len(infos)):
        info = infos[i]
        rId = info['rId']
        text = rId if 'text' not in info else info['text']
        off_y = 0.48 if 'off_y' not in info else info['off_y']
        para += p.h4('（%d）                   %s数据库证据呈现' % (i + 1, text), runs=r.picture(cy=1, rId=info['rId'], posOffset=[1.3, off_y]))
        for p_text in info['para']:
            para += p.write(r.text(p_text['text']))
            if 'pic' in p_text:
                pic = p_text['pic']
                para += p.write(p.set(spacing=[9, 0]), r.picture(cy=pic['cy'], rId=pic['rId'], posOffset=[0, 5.59], align=['center', 'bottom']))
    return para


def write_gene_info(gene_item, index):
    gene = gene_item['gene']
    para = p.h4('（%d）基因说明:%s，%s' % (index, gene, gene_item['gene_MoA']))
    for item in gene_list12:
        if item['hugoSymbol'] == gene:
            if 'geneAliases' in item:
                if len(item['geneAliases']) > 0:
                    para += p.write(r.text('Also known as %s' % ', '.join(item['geneAliases'])))
            text = ''
            if 'curatedIsoform' in item:
                text += 'Isoform: %s' % item['curatedIsoform']
            if 'curatedRefSeq' in item:
                text += 'Isoform: %s' % item['curatedRefSeq']
            if len(text) > 0:
                para += p.write(r.text(text))
            description = item['description']
            summary = item['description']
            if len(summary) > 0:
                para += p.write(r.text(summary))
            if len(description) > 0:
                para += p.write(r.text(description))
    return para


def write_patient_info(data):
    para = p.h4('基本信息')
    trs = ''
    ws = [1700, 1800, 2700, 1800]
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
            tcs2 += tc.write(p.write(pPr, run), tc.set(w=ws[i], tcBorders=[], fill=gray))
        trs += tr.write(tcs2)
    para += table.write(trs, ws=ws, tblBorders=[])
    return para


# part1 靶向治疗提示
def write_target_tip(title, items):
    n = len(title)
    ws = [9000 / n] * n
    pPr = p.set(jc='left', spacing=[0.5, 0.5])
    trs = write_thead51(title, ws=ws)
    for k in range(len(items)):
        tcs2 = ''
        size = 9
        item1 = items[k]
        item = [item1['col1']]
        for level in level_names:
            item.append(item1[level])
        if len(item) < n:
            item.append(item1['gene_MoA'])
        for i in range(n):
            texts = item[i]
            para = ''
            for t in texts.split('\n'):
                run = r.text(t, size=size)
                para += p.write(pPr, run)
            tcs2 += tc.write(para, tc.set(w=ws[i], color=gray))
        trs += tr.write(tcs2)
    return table.write(trs, ws=ws, tblBorders=[])


def write_immun_tip():
    para = p.h4('免疫治疗提示') + con1 + con2
    para += write_immun_tip1() + p.write(p.set(line=2, rule='exact'), r.br())
    # para += con2
    para += write_immun_tip2() + con2
    return para


def write_immun(data, jc, h0=900, w=2800):
    trs2 = write_tr1(data[0]) + write_tr2(data[1])
    table_str = table.write(trs2, ws=[w], jc=jc, bdColor=blue)
    return table_str


def write_blue_table(paras, jc, h0=900, w=2800):
    trs2 = ''
    for k in range(len(paras)):
        color, fill, h = blue, 'auto', 800
        if k == 0:
            color, fill, h = 'FFFFFF',  blue, h0
        tcs2 = tc.write(paras[k], tc.set(w=w, fill=fill, color=blue))
        trs2 += tr.write(tcs2, tr.set(trHeight=h))
    table_str = table.write(trs2, ws=[w], jc=jc, bdColor=blue)
    return table_str


def write_immun_tip1(jc='center'):
    data1 = get_immu('MSI')[1:]
    return write_immun(data1, jc, w=3600)


def write_immun_tip2(jc='center'):
    data2 = get_immu('TMB')[1:]
    return write_immun(data2, jc, w=3200)


def write_evidence4(index):
    para = p.write(p.set(spacing=[0.5, 1]), r.text('相关循证医学证据：'))
    evidence = evidence4[index]
    for i in range(len(evidence)):
        if i % 4 == 0:
            para += p.write(p.set(spacing=[0, 0.2], shade=gray), r.text(evidence[i], 9.5))
        if i % 4 == 2:
            para += p.write(p.set(spacing=[0, 1], shade=gray), r.text(evidence[i], 9.5))
    return para


def write_evidence(data, **kwargs):
    trs = ''
    ws = [2000] * 4
    pPr = p.set(jc='left', spacing=[0.5, 0.5])
    if 'titles' not in kwargs:
        titles = ['', '生物标志物', '药物', '证据类型', '匹配程度', '证据来源']
    else:
        titles = kwargs['titles']
    data.insert(0, titles)
    for k in range(len(data)):
        tcs = ''
        size = 9
        item = data[k][1:]
        for i in range(4):
            t = item[i]
            if isinstance(t, list):
                t = '; '.join(['%s(%s级)' % (tt['drug'], tt['level']) for tt in t])
            run = r.text(t, size=size)
            fill = gray if k == 0 else 'auto'
            tcs += tc.write(p.write(pPr, run), tc.set(w=ws[i], tcBorders=[], fill=fill))
        if len(tcs) > 0:
            trs += tr.write(tcs)
            if k > 0:
                # print len(item)
                run = r.text(titles[4], size=size)
                tc5 = tc.write(p.write(pPr, run), tc.set(w=ws[0], tcBorders=[], fill=gray))
                run1 = r.text(item[4], size=size)
                tc5 += tc.write(p.write(pPr, run1), tc.set(w=6000, tcBorders=[], fill='auto', gridSpan=3))
                trs += tr.write(tc5)
    return table.write(trs, ws=ws, tblBorders=[])


def write_explain(i, ind=[0, 0]):
    run = r.text(i['title'], weight=1) + r.text(i['text'], size=9)
    return p.write(p.set(ind=ind), run)


def write_chemotherapy(trs, ws):
    trs2 = ''
    borders=['top', 'bottom', 'left', 'right']
    for k in range(len(trs)):
        size = 10
        items = trs[k]
        tcs = ''
        for j in range(len(items)):
            item = items[j]
            if isinstance(item, list):
                item = ';'.join(uniq_list(item))
                if item == '':
                    item = '无'
            fill, weight, jc = 'auto', 0, 'left'
            if k == 0 or j == 0:
                fill, weight, jc = gray, 1, 'center'
            pPr = p.set(jc=jc, spacing=[0.5, 0.5])
            run = r.text(item, size=size, weight=weight)
            tcs += tc.write(p.write(pPr, run), tc.set(w=ws[k], fill=fill, color=gray_lighter, tcBorders=borders))
        trs2 += tr.write(tcs)
    table_str = table.write(trs2, ws=ws, bdColor=gray)
    return table_str


def write_gene_list3(genes, width=9000):
    trs2 = ''
    fill, weight, jc = gray, 0, 'center'
    pPr = p.set(jc=jc, line=12, rule='auto', spacing=[0.5, 0.5])
    col = 5
    ws = [width / col] * col
    row = int(math.ceil(float(len(genes)) / col))
    for k in range(row):
        tcs = ''
        size = 9
        for j in range(col):
            this_index = k * col + j
            if this_index < len(genes):
                item = genes[this_index]
                gene = item['gene']
                para = p.write(pPr, r.text('%s %s(%s)' % (gene, item['rs'], item['genotype']), size))
                para += p.write(pPr, r.text(item['summary'], size))
                tcs += tc.write(para, tc.set(w=ws[k], fill=fill, color=gray_lighter, tcBorders=[], line_type='double'))
        trs2 += tr.write(tcs)
    table_str = table.write(trs2, ws=ws, bdColor=blue, tblBorders=[], border_size=8)
    return table_str


def write_genotype(gts, ws):
    trs2 = ''
    borders = ['top', 'bottom', 'left', 'right']
    for j in range(len(gts)):
        size = 9
        tcs = ''
        gt = gts[j]
        fill, weight, jc = 'auto', 0, 'left'
        if j == 0:
            gridSpan = 2
            text = '%s %s (%s)' % (gt['gene'], gt['rs'], '证据级别%s' % gt['level'])
            trs2 += tr.write(tc.write(p.write(p.set(jc='center', spacing=[0.5, 0.5]), r.text(text, size=size, weight=1)), tc.set(w=sum(ws), fill=gray, color=gray_lighter, tcBorders=borders, gridSpan=gridSpan)))
        for k in range(2):
            key = 'introduction'
            fill = 'auto'
            if k == 0:
                fill, jc, key = gray, 'center', 'genotype'
            pPr = p.set(jc=jc, spacing=[0.5, 0.5])
            run = r.text(gt[key], size=size, weight=weight)
            tcs += tc.write(p.write(pPr, run), tc.set(w=ws[k], fill=fill, color=gray_lighter, tcBorders=borders))
        trs2 += tr.write(tcs)
    table_str = table.write(trs2, ws=ws, bdColor=gray_lighter)
    return table_str + p.write()


def write_41():
    genes = [
        {
            'title': '同源重组修复基因', 'color': colors[0],
            'genes': ['BRCA1', 'MRE11A', 'NBN', 'RAD50', 'RAD51', 'RAD51B', 'RAD51D', 'RAD52', 'RAD54L']
         },
        {
            'title': '范可尼贫血通路基因', 'color': colors[1],
            'genes': ['BRCA2', 'BRIP1', 'FANCA', 'FANCC', 'PALB2', 'RAD51C', 'BLM']
        },
        {
            'title': '碱基切除修复基因', 'color': colors[2],
            'genes': ['ERCC2', 'ERCC3', 'ERCC4', 'ERCC5', 'POLE4', 'POLD1']
        },
        {
            'title': '错配修复基因', 'color': colors[3],
            'genes': ['MLH1', 'MLH2', 'MSH6', 'PMS1', 'PMS2']
        },
        {
            'title': '细胞周期检查点基因', 'color': colors[4],
            'genes': ['ATM', 'ATR', 'CHEK1', 'CHEK2', 'MDC1']
        },
        {
            'title': '其他基因', 'color': colors[5],
            'genes': ['POLE', 'MUTYH', 'PARP1', 'RECQL4']
        },
        {
            'title': '染色质重塑基因', 'color': colors[6],
            'genes': ['PBRM1', 'ARID1', 'BRD7']
        }
    ]
    return write_genes4(genes, 10, 10000)


def write_43(genes, **kwargs):
    col = len(genes)
    ws = [1800, 1800, 1800, 1800, 2400] if 'ws' not in kwargs else kwargs['ws']
    fill, weight, jc = gray, 0, 'center'
    pPr = p.set(jc=jc, line=12, rule='auto')
    tcs = ''
    for j in range(col):
        item = genes[j]
        fill = item['fill']
        color = '000000'
        if fill != gray:
            color = white
        para = p.write(pPr, r.text(item['text'], color=color, size=9))
        tcs += tc.write(para, tc.set(w=ws[j], fill=fill, tcBorders=[], line_type='double'))
    trs2 = tr.write(tcs, tr.set(trHeight=800))
    return table.write(trs2, ws=ws, tblBorders=[])


def write_genes(gene_list, col, whith, table_jc='center'):
    trs2 = ''
    ws = [whith/col] * col
    fill, weight, jc = gray, 0, 'center'
    pPr = p.set(jc=jc, line=12, rule='auto')
    row = int(math.ceil(float(len(gene_list)) / 5))
    for i in range(row):
        tcs = ''
        for j in range(col):
            gene_index = col * i + j
            if gene_index < len(gene_list):
                item = gene_list[gene_index]
                fill = get_var(item)
                # fill = gray
                color, text, var_text = '000000', item, ''
                if fill != gray:
                    color = white
                    var_text = p.write(pPr, r.text('突变', color=color, size=9))
                para = p.write(pPr, r.text(text, color=color, size=9)) + var_text
                tcs += tc.write(para, tc.set(w=ws[j], fill=fill, line_type='double'))
        trs2 += tr.write(tcs, tr.set(trHeight=580))
    table_str = table.write(trs2, ws=ws, tblBorders=[], jc=table_jc)
    return table_str


def write_genes4(genes, col, whith, table_jc='center'):
    trs2 = ''
    ws = [whith/col] * col
    fill, weight, jc = gray, 0, 'center'
    pPr = p.set(jc=jc, line=12, rule='auto')
    reds = 0
    oranges = 0
    for k in range(len(genes)):
        gene = genes[k]
        gene_list = gene['genes']
        tcs = ''
        if 'title' in gene:
            para = p.write(pPr, r.text(gene['title'], color=white, size=9))
            tcs += tc.write(para, tc.set(w=ws[k], fill=gene['color'], line_type='double'))
        for j in range(len(gene_list)):
            item = gene_list[j]
            fill = get_var(item)
            # fill = gray
            color, text, var_text = '000000', item, ''
            if fill != gray:
                color = white
                if fill == red:
                    reds += 1
                else:
                    oranges += 1
                var_text = p.write(pPr, r.text('突变', color=color, size=9))
            para = p.write(pPr, r.text(text, color=color, size=9)) + var_text
            tcs += tc.write(para, tc.set(w=ws[j], fill=fill, line_type='double'))
        trs2 += tr.write(tcs, tr.set(trHeight=580))
    table_str = table.write(trs2, ws=ws, tblBorders=[], jc=table_jc)
    return table_str, reds, oranges


def write_thead51(titles, **kwargs):
    ws = [1800, 1600, 1600, 1600, 1800]
    pPr = p.set(jc='left', spacing=[0.5, 0.5])
    tcBorders = ['top', 'bottom', 'left', 'right']
    if 'pPr' in kwargs:
        pPr = kwargs['pPr']
    if 'ws' in kwargs:
        ws = kwargs['ws']
    if 'tcBorders' in kwargs:
        tcBorders = kwargs['tcBorders']
    tcs = ''
    size = 10
    for i in range(len(titles)):
        t = titles[i]
        run = r.text(t, size=size, weight=0)
        tcs += tc.write(p.write(pPr, run), tc.set(w=ws[i], tcBorders=tcBorders, fill=gray, color=gray_lighter))
    return tr.write(tcs)


def write_tr1(text, jc='center', w=4000):
    pPr = p.set(jc=jc, spacing=[0.5, 0.5])
    para = ''
    for t in text.split('\n'):
        run = r.text(t, size=9, color=white)
        para += p.write(pPr, run)
    tcs2 = tc.write(para, tc.set(w=w, fill=blue, color=blue))
    return tr.write(tcs2)


def write_tr2(data, jc='center', w=4000):
    pPr = p.set(jc=jc, spacing=[1, 1])
    para = ''
    for t in data.split('\n'):
        para += p.write(pPr, r.text(t, size=10, color=blue))
    tcs2 = tc.write(para, tc.set(w=w, fill=white, color=blue))
    return tr.write(tcs2)
