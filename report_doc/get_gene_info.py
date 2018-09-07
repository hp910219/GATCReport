#! /usr/bin/env python
# coding: utf-8
__author__ = 'huo'
import requests

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

