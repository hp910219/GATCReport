#! /usr/bin/env python
# coding: utf-8
__author__ = 'huo'
import requests
import base64
from bs4 import BeautifulSoup
from config import data_dir
import os
import sys
reload(sys)
sys.setdefaultencoding('utf-8')


def get_gene_info(gene):
    summary = ''
    description = ''
    url1 = 'http://oncokb.org/api/v1/evidences/lookup?hugoSymbol=%s&evidenceTypes=GENE_SUMMARY' % gene
    url2 = 'http://oncokb.org/api/v1/evidences/lookup?hugoSymbol=%s&evidenceTypes=GENE_BACKGROUND' % gene
    data1 = my_request(url1)
    data2 = my_request(url2)
    item = {'hugoSymbol': gene}
    if data2 is not None:
        if len(data2) > 0:
            info = data2[0]
            if 'gene' in info:
                gene_info = info['gene']
                item.update(gene_info)
            if 'description' in info:
                description = info['description']
    if data1 is not None:
        if len(data1) > 0:
            if 'description' in data1[0]:
                summary = data1[0]['description']
    item['summary'] = summary
    item['description'] = description
    return item


def get_evidence():
    items = []
    gene_items = []
    gene_list = []
    for level in ['1', '2A', '3A', '4']:
        url1 = 'http://oncokb.org/legacy-api/evidence.json?levels=LEVEL_%s' % level
        data1 = my_request(url1)
        if data1 is not None:
            print level, len(data1[0])
            items += data1[0]
            for item in data1[0]:
                gene = item['gene']['hugoSymbol']
                if gene not in gene_list:
                    gene_list.append(gene)
                    gene_items.append(get_gene_info(gene))
    return items


def my_request(url, r='json'):
    try:
        rq = requests.get(url)
    except:
        return None
    if rq.status_code == 200:
        if r == 'json':
            data = rq.json()
            return data
        return rq.text
    return None


def get_pcgr(url, localpath):
    htmlfile = open(url, 'r')  #以只读的方式打开本地html文件
    htmlpage = htmlfile.read()
    htmlfile.close()
    soup = BeautifulSoup(htmlpage, "html.parser")  #实例化一个BeautifulSoup对象
    if os.path.exists(localpath) == False:                  #判断，若该文件路径不存在，则创建该目录（mkdirs创建多级目录，midir创建单级目录）
        os.makedirs(localpath)
    snv_num, indel_num, var_num = 0, 0, 0
    somatic = soup.select('#tier-variant-statistics')
    if len(somatic) > 0:
        somatic1 = somatic[0]
        ul = somatic1.select('ul')
        if len(ul) > 0:
            li = ul[0].select('li')
            for ele in li:
                text = ele.text.strip().strip('\n').strip()
                strong = ele.select('strong')
                if len(strong) > 0:
                    if 'Number of SNVs: ' in text:
                        snv_num = strong[0].text.strip().strip('\n').strip()
                    if 'Number of InDels: ' in text:
                        indel_num = strong[0].text.strip().strip('\n').strip()
                    if 'Number of protein-coding variants: ' in text:
                        var_num = strong[0].text.strip().strip('\n').strip()
    msi = soup.select('#supporting-evidence-indel-fraction-among-somatic-calls img')
    signatures = soup.select('#mutational-signatures')
    if len(msi) > 0:
        save_img('%s/msi.png' % localpath, msi[0])
    items = []
    if len(signatures) > 0:
        sig = signatures[0]
        script = sig.select('script')
        if len(script) > 0:
            import json
            data = json.loads(script[0].text)
            keys = ['', 'Signature_ID', 'Weight', 'Cancer_types', 'Proposed_aetiology', 'Comments', 'Keyword', 'Trimer_normalization_method', 'Sample_Name']
            if 'x' in data:
                if 'data' in data['x']:
                    maindata = data['x']['data']
                    if len(maindata) > 0:
                        num = len(maindata[0])
                        for i in range(num):
                            item = {}
                            for k in range(len(keys)):
                                key = keys[k]
                                item[key] = maindata[k][i]
                            items.append(item)
        sigs = sig.find_all('img')
        if len(sigs) > 0:
            save_img('%s/signature.png' % localpath, sigs[0], img2='4.5.1signature.png')
        if len(sigs) > 1:
            save_img('%s/signature_pie.png' % localpath, sigs[1], img2='4.5.2signature_pie.png')
    return {"signature_etiology": items, 'snv_num': snv_num, 'indel_num': indel_num, 'var_num': var_num}


def save_img(url, i, img2=None):
    if os.path.exists(url) is False:
        img = base64.b64decode(i.attrs['src'].split(',')[1])
        f = open(url, 'wb')
        f.write(img)
        f.close()
        print 'save img success: ', url
    else:
        print 'img exists: ', url
    # if img2 is not None:
    #     crop_img(url, r'%s/%s.png' % (data_dir, img2))
