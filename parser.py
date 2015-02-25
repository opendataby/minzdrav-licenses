"""
Script parse doc data from
http://minzdrav.gov.by/ru/static/licensing/reestr_licences
and present csv file with parsed data.
Data in csv is same as in doc without splitting, normalizing or something
else.

os requirements:

    vw
    unrar
    libxml

python requirements:

    lxml
    requests
    rarfile

"""

from collections import OrderedDict
import csv
import json
import os
import unittest
import sys

from lxml import etree
import rarfile
import requests


URL = 'http://minzdrav.gov.by/ru/static/licensing/reestr_licences'


DOC_PATH = 'docs'
TEST_PATH = 'test'
DOC_EXT = '.doc'
HTML_EXT = '.html'
CHECK_EXT = '.json'


ST_COMPANY = 'company'
ST_ITEMS = 'items'
ST_PROPS = 'properties'


_fieldnames = (
    'category',
    'extra_region',
    'company_name',
    'company_addr',
    'company_license_number',
    'company_license_start_date',
    'company_license_end_date',
    'office_name',
    'office_addr',
    'properties',
)


_replaces = (
    ('<i>\t</i>', '\t'),
)


_captions = {
    'Брестская область': 'brest',
    'Витебская область': 'vitebsk',
    'Гомельская область': 'gomel',
    'Гродненская область': 'grodno',
    'Минская область': 'minobl',
    'Могилевская область': 'mogilev',
    'г.Минск': 'minsk',
}


_categories = OrderedDict((
    ('psihologia', 'psychology'),
    ('narko_', 'drugs'),
    ('farm_', 'pharmacy'),
    ('', 'medicine'),
))


def unrar(zipped_file_name):
    archive = rarfile.RarFile(zipped_file_name)
    info = archive.infolist()
    assert len(info) == 1
    file_name = info[0].filename
    archive.extract(info[0])
    return file_name


def fetch_docs():
    page = etree.HTML(requests.get(URL).content)
    for link in page.xpath('.//div[@id="content"]//a'):
        url = link.attrib['href']
        file_name = url.rsplit('/', 1)[-1]
        with open(file_name, 'wb') as file:
            file.write(requests.get(url).content)
        yield file_name


def process_text(file, extra_region=None):
    state = ST_ITEMS
    skip = False
    companies = []
    company = None
    items = None
    item = None
    properties = None

    for event, element in etree.iterparse(file, tag=('p', 'table'),
                                          events=('start', 'end'), html=True):
        # skip document header legend
        if event == 'start':
            if element.tag == 'table':
                skip = True
            continue
        if event == 'end' and element.tag == 'table':
            skip = False
            continue
        if skip:
            continue

        # prepare row
        company_caption = bool(len(element.xpath('.//b')) >= 5)
        formatted_property = bool(len(element.xpath('.//i')) >= 1)

        xml = ''.join([element.text or ''] + [etree.tostring(sub, encoding='utf8').decode()
                                              for sub in element.getchildren()])
        for replace_from, replace_to in _replaces:
            xml = xml.replace(replace_from, replace_to)
        if not xml.strip():
            continue
        subelements = [etree.fromstring('<span>' + sub.strip() + '</span>')
                       for sub in xml.split('\t')]
        caption = ' '.join([''.join([x for x in sub.itertext()]).strip()
                            for sub in subelements
                            if sub.xpath('.//font/u/b') or sub.xpath('.//font/b')]).strip()
        line = ' '.join([''.join([x for x in sub.itertext()]).strip()
                         for sub in subelements
                         if not sub.xpath('.//font/u/b') and not sub.xpath('.//font/b')]).strip()
        if caption in _captions:
            extra_region = caption
            continue

        # populate caption data
        if company_caption and state != ST_COMPANY:
            state = ST_COMPANY
            company = [''] * 5
            items = []
            companies.append([company, extra_region, items])
        for index, subelement in enumerate(subelements):
            if subelement.xpath('.//font/u/b') or subelement.xpath('.//font/b'):
                company[index] += ' ' + ''.join([x for x in subelement.itertext()]).strip()

        # populate offices and properties
        if not line:
            continue
        if not formatted_property:
            if state != ST_ITEMS:
                state = ST_ITEMS
                properties = []
                if ' - ' in line:
                    item = line.split(' - ', 1) + [properties]
                else:
                    item = ['', line, properties]
                items.append(item)
            else:
                if not item[0]:
                    item[1] += ' ' + line
                    if ' - ' in item[1]:
                        item[0], item[1] = item[1].split(' - ', 1)
                else:
                    item[1] += ' ' + line
            continue
        else:
            state = ST_PROPS
            properties.append(line)
            continue
    return companies


def do():
    writer = csv.DictWriter(open('raw_med.csv', 'w'), fieldnames=_fieldnames)
    writer.writeheader()

    for zipped_file in fetch_docs():
        doc_file = unrar(zipped_file)
        html_file = doc_file[:-len(DOC_EXT)] + HTML_EXT
        os.system('wvHtml {} {}'.format(doc_file, html_file))

        extra_region = None
        for name, pattern in _captions.items():
            if pattern in html_file:
                extra_region = name
                break

        data = process_text(open(html_file, 'rb'), extra_region)

        for name, category in _categories.items():
            if doc_file.startswith(name):
                break
        else:
            raise TypeError

        for company, extra_region, items in data:
            assert len(company) == 5
            assert items
            assert extra_region
            company_name = company[0].strip()
            company_addr = company[1].strip()
            company_license_number = company[2].strip()
            company_license_start_date = company[3].strip()
            company_license_end_date = company[4].strip()
            for office_name, office_addr, properties in items:
                assert properties
                writer.writerow({
                    'category': category,
                    'extra_region': extra_region.strip(),
                    'company_name': company_name.strip(),
                    'company_addr': company_addr.strip(),
                    'company_license_number': company_license_number.strip(),
                    'company_license_start_date': company_license_start_date.strip(),
                    'company_license_end_date': company_license_end_date.strip(),
                    'office_name': office_name.strip(),
                    'office_addr': office_addr.strip(),
                    'properties': '\t'.join(properties).strip(),
                })

        os.remove(zipped_file)
        os.remove(doc_file)
        os.remove(html_file)


class Test(unittest.TestCase):

    def test_patterns(self):
        for file in os.listdir(TEST_PATH):
            if not file.endswith(HTML_EXT):
                continue
            html_path = os.path.join(TEST_PATH, file)
            test_path = html_path[:-len(HTML_EXT)] + CHECK_EXT
            self.assertEqual(process_text(open(html_path, 'rb')),
                             json.load(open(test_path)), file)


if __name__ == '__main__':
    if len(sys.argv) == 2 and sys.argv[1] == '--test':
        unittest.main()
    else:
        do()
