# -*- coding: utf-8 -*-
# Date: 2015/03/
# Author: lms

import requests
import re
import json
import xlwt
import time
from lxml import html


def crawl_key_item(table):
    static_url = 'http://summary.jrj.com.cn/scfl/ssbg.shtml?q=cn|s|sb,sz&c=m&n=hqa&o=pl,d&p=1020'

    page = requests.get(static_url)
    tree = html.fromstring(page.text)

    # crawl the key items
    key_item = [tree.xpath('//th[@class="w5 first"]/span/text()'),
                tree.xpath('//th[@class="w5"]/span/text()'),
                tree.xpath('//th[@name="np"]/span/text()'),
                tree.xpath('//th[@name="pl"]/span/text()'),
                tree.xpath('//th[@name="ta"]/span/text()'),
                tree.xpath('//th[@name="tm"]/span/text()'),
                tree.xpath('//th[@name="hp"]/text()'),
                tree.xpath('//th[@name="lcp"]/text()'),
                tree.xpath('//th[@name="ape"]/span/text()')]

    table['0'] = key_item


def crawl_val_item(table):
    date_url_list = ['http://q.jrjimg.cn/?q=cn|s|sb,sz&c=m&n=hqa&o=pl,d&p=1020',
                     'http://q.jrjimg.cn/?q=cn|s|sb,sz&c=m&n=hqa&o=pl,d&p=2020',
                     'http://q.jrjimg.cn/?q=cn|s|sb,sz&c=m&n=hqa&o=pl,d&p=3020']

    # each page contains only 20 records
    gap = 0
    for url in date_url_list:
        r = requests.get(url)

        # use regular expression to find the specific information
        data = re.findall('"sz[^\]]*', r.text)

        for i in range(len(data)):
            record = data[i].split(',')

            # store the data of each record
            table[str(i + gap + 1)] = [record[1].replace('\"', ''),
                                       record[2].replace('\"', ''),
                                       record[8],
                                       record[12], record[9], record[10],
                                       record[6] + '/' + record[7],
                                       record[3] + '/' + record[5],
                                       record[21]]
        gap += 20


def generate_xls(table):
    tabledata = json.loads(json.dumps(table))

    #add a xls
    book = xlwt.Workbook(encoding='utf-8')
    sheet = book.add_sheet('stock')

    for row in range(len(tabledata)):
        rowdata = tabledata[str(row)]

        for col in range(len(rowdata)):
            sheet.write(row, col, rowdata[col])
    book.save('stock.xls')


if __name__ == '__main__':
    while True:
        table = {'0': 0}
        crawl_key_item(table)
        crawl_val_item(table)
        generate_xls(table)
        time.sleep(600)