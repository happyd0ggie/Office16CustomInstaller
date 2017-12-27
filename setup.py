#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
import os
import sys
import xml.etree.ElementTree as ET

class Setup(object):
    def __init__(self, args):
        self.args = args
        self.config_file = 'configuration.xml'
        self.lang = self._get_product_lang()
        self.all_products = ['Word', 'Excel', 'PowerPoint', 'Access',
            'Groove', 'InfoPath', 'Lync', 'OneNote', 'Project',
            'Publisher', 'Visio', 'SharePointDesigner', 'OneDrive']
        self.product_list_to_install = self._get_product_list()
        self.product_edition = self._get_product_edition()
        self._init_config_file()
        self._gen_config_file()

    def _get_product_list(self):
        product_list = []
        if ',' in self.args.product:
            product_list.extend(self.args.product.split(','))
        else:
            product_list.append(self.args.product)

        return product_list

    def _get_product_edition(self):
        return self.args.edition

    def _get_product_lang(self):
        return self.args.lang

    def _init_config_file(self):
        init_xml_str1 = '<Configuration>'
        init_xml_str2 = '    <Add SourcePath="Office" Branch="Current" OfficeClientEdition="64">'
        init_xml_str3 = '        <Product ID="ProPlusRetail">'
        init_xml_str4 = '            <Language ID="zh-cn" />'
        init_xml_str5 = '        </Product>'
        init_xml_str6 = '    </Add>'
        init_xml_str7 = '</Configuration>'
        
        if os.path.exists(self.config_file):
            os.unlink(self.config_file)

        with open(self.config_file, 'w') as f:
            f.write(init_xml_str1 + '\n')
            f.write(init_xml_str2 + '\n')
            f.write(init_xml_str3 + '\n')
            f.write(init_xml_str4 + '\n')
            f.write(init_xml_str5 + '\n')
            f.write(init_xml_str6 + '\n')
            f.write(init_xml_str7 + '\n')

    def _gen_config_file(self):
        tree = ET.parse(self.config_file)
        root = tree.getroot()

        # update product edition
        root[0].set('OfficeClientEdition', self.product_edition)

        # update product language
        for lang in root.iter('Language'):
            lang.set('ID', self.lang)

        # update product that will not be installed
        for product in root.iter('Product'):
            for item in self.all_products:
                if item not in self.product_list_to_install:
                    app = ET.SubElement(product, 'ExcludeApp')
                    app.set('ID', item)
        ET.dump(root)
        tree.write(self.config_file)

    def run(self):
        if self.args.action == 'download':
            os.system('.\setup.exe /download {0}'.format(self.config_file))
        elif self.args.action == 'install':
            os.system('.\setup.exe /configure {0}'.format(self.config_file))
        else:
            pass

def get_argparser():
    desc = '''Microsoft Office 2016 downloader/installer\n
        example:\n
            python setup.py --action install --product word --edition 64 --lang zh-cn
    '''
    parser = argparse.ArgumentParser(prog=sys.argv[0], description=desc)
    parser.add_argument('-k', dest='action', help='install | download')
    parser.add_argument('-p', dest='product', help='product to install, e.g. Word/PowerPoint/Excel/Outlook/Access/InfoPath/Groove/Lync/OneDrive/Published/Project/Visio/SharePointDesigner')
    parser.add_argument('-e', dest='edition', help='product edition, e.g. 64/32')
    parser.add_argument('-l', dest='lang', help='install language, e.g. en-us/zh-cn')
    return parser

def main():
    parser = get_argparser()
    args = parser.parse_args()

    if (args.action == None or args.product == None or 
        args.lang == None or args.edition == None):
        parser.print_usage()
        sys.exit(1)

    setup = Setup(args)
    setup.run()    

if __name__ == '__main__':
    main()

