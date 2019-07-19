# -*- coding: utf-8 -*-

import codecs
import csv
import glob
import os
import sys
from bs4 import BeautifulSoup
from collections import OrderedDict

import xlsxwriter
from creole import creole2html

DATA_DIR = "./data"
OUTPUT_DIR = "./output"


def scan_for_data(data_path):
    return glob.glob(data_path + '/*.txt')


def process_file(filepath):
    with codecs.open(filepath, 'r', 'utf-8') as fd:
        parsed = parse_content(fd.read())
    return parsed


def parse_content(content):
    html_content = creole2html(content)
    soup = BeautifulSoup(html_content, features="html.parser")
    tables = soup.find_all("table")
    output_rows = []
    labels = []
    table_type_names = []
    for table in tables:

        try:
            table_type = table.findPrevious("h4").text
        except AttributeError:
            table_type = "_"
        #print "table nameee typee:--->", table_type
        y = 1
        while table_type in table_type_names:
            table_type = "{table_type}_{counter}".format(table_type=table_type,
                                                         counter=y)
            y += 1
        #print "TABLE_name:", table_type_names, table_type
        table_type_names.append(table_type)

        for table_row in table.findAll('tr'):
            columns = table_row.findAll('td')
            if len(columns) == 0:
                continue
            if len(columns) == 1:
                label = columns[0].text
                value = ""
            elif len(columns) > 2:
                label = columns[0].text
                value = ' '.join([c.text for c in columns[1:]])
            else:
                label, value = columns
                label = label.text
                value = value.text
            x = 1
            while label in labels:
                label = u"{label}_{counter}".format(label=label, counter=x)
                x += 1
            label = u"{label} [{type}]".format(label=label, type=table_type)
            labels.append(label)
            output_rows.append((unicode(label
                                        .replace(':', '')
                                        .strip()
                                        .capitalize()), unicode(value.strip())))
    return OrderedDict(output_rows)


def process_file2csv(filepath):
    basename = os.path.basename(filepath)
    out_path = os.path.join(OUTPUT_DIR, basename + '.csv')
    with codecs.open(filepath, 'r', 'utf-8') as fd:
        parsed = parse_content(fd.read())
        with open(out_path, 'wb') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerows(parsed)


0


def main():
    files = scan_for_data(DATA_DIR)
    #files = ['./data/smolnicki_tomasz.txt']
    files_count = len(files)
    print >> sys.stderr, 'FILES NUMBER:', files_count
    big_data = OrderedDict()
    workbook = xlsxwriter.Workbook('data.xlsx')
    worksheet = workbook.add_worksheet()

    header_counter = OrderedDict()

    print 'PARSE DATA FROM FILES:',
    cnt = 0
    big_data['FILE ID'] = []
    for f in files:
        print f
        parsed_data = process_file(f)
        big_data['FILE ID'].append(f)
        tmp_labels_list = ['FILE ID']
        for item in parsed_data:
            if item not in big_data:
                header_counter[item] = 1
                print u"NEW COL:;  {name:30} ;[{file}] ;({cnt})".format(name=item,
                                                                        file=f,
                                                                        cnt=cnt)
                if cnt == 0:
                    big_data[item] = []
                else:
                    empty_list = [""] * cnt
                    big_data[item] = empty_list
            header_counter[item] += 1
            tmp_labels_list.append(item)
            big_data[item].append(parsed_data[item])
        all_lables = set(big_data.keys())
        labels_diff = all_lables.difference(set(tmp_labels_list))
        for lb in labels_diff:
            big_data[lb].append("")

        cnt += 1
        # print "\r\t\t{}/{} {:04.2f}%".format(cnt, files_count, (float(cnt) / files_count) * 100),

    pos = 0
    for l in big_data:
        to_write = [l]
        to_write.extend(big_data[l])
        worksheet.write_column(0, pos, to_write)
        pos += 1
    workbook.close()

    # print header_counter


if __name__ == '__main__':
    main()
