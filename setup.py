#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
This script enable custom installation of Microsoft Office suite.
You can install/uninstall specific product.
"""

import argparse
import os
import sys
import xml.etree.ElementTree as ET

ALL_PRODUCTS = ['Word', 'Excel', 'PowerPoint', 'Access',
                'Groove', 'InfoPath', 'Lync', 'OneNote', 'Project', 'Outlook',
                'Publisher', 'Visio', 'SharePointDesigner', 'OneDrive']

class Setup(object):
    """Microsoft Office 2016 custom installer wrapper."""
    def __init__(self, args):
        self.args = args
        self.config_file = 'configuration.xml'
        self.lang = self._get_product_lang()
        self.all_products = ALL_PRODUCTS
        self.product_list_to_install = self._get_product_list()
        self.product_edition = self._get_product_edition()
        self._init_config_file()
        self._gen_config_file()

    def _get_product_list(self):
        """Get products to be installed/uninstalled..

        Get products to be installed/uninstalled.

        Args:
            None.

        Returns:
            None.

        """
        product_list = []
        if ',' in self.args.product:
            product_list.extend(self.args.product.split(','))
        else:
            product_list.append(self.args.product)

        return product_list

    def _get_product_edition(self):
        """Get product edition to be used.

        Get product edition to be used.

        Args:
            None.

        Returns:
            None.

        """
        return self.args.edition

    def _get_product_lang(self):
        """Get product language to be used.

        Get product language to be used.

        Args:
            None.

        Returns:
            None.

        """
        return self.args.lang

    def _init_config_file(self):
        """Initialize configuration file.

        Initialize configuration file template.

        Args:
            None.

        Returns:
            None.
        """
        print('Initializing Configuration File'.center(60, '='))
        init_xml_str = ('<Configuration>\n'
                        '    <Add SourcePath="Office" Branch="Current" OfficeClientEdition="64">\n'
                        '        <Product ID="ProPlusRetail">\n'
                        '            <Language ID="zh-cn" />\n'
                        '        </Product>\n'
                        '    </Add>\n'
                        '</Configuration>\n')

        if os.path.exists(self.config_file):
            os.unlink(self.config_file)

        open(self.config_file, 'w').write(init_xml_str + '\n')

    def _gen_config_file(self):
        """Generate configuration file.

        Generate configuration file, which will be used to custom
        installation.

        Args:
            None.

        Returns:
            None.
        """
        print('Generating Configuration File'.center(60, '='))
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
        """Class entry point.

        Args:
            None.

        Returns:
            None.
        """
        if self.args.action == 'download':
            os.system('.\setup.exe /download {0}'.format(self.config_file))
        elif self.args.action == 'install':
            os.system('.\setup.exe /configure {0}'.format(self.config_file))
        else:
            pass
        os.unlink(self.config_file)

def get_argparser():
    """Generate a command line argument parser."""
    parser = argparse.ArgumentParser(
        prog=sys.argv[0],
        description='Microsoft Office 2016 downloader/installer',
        epilog=('e.g.: python setup.py --action install --product word '
                '--edition 64 --lang zh-cn'))
    parser.add_argument('-a', '--action', action='store',
                        default='install', help='install | download')
    parser.add_argument('-p', '--product', action='store', required=True,
                        choices=ALL_PRODUCTS, help='product to install')
    parser.add_argument('-e', '--edition', action='store', default='64',
                        help='product edition, e.g. 64/32')
    parser.add_argument('-l', '--lang', action='store', default='zh-cn',
                        help='install language, e.g. en-us/zh-cn')
    return parser

def main():
    """Program entry point."""
    parser = get_argparser()
    args = parser.parse_args()

    if (args.action is None or args.product is None or
            args.lang is None or args.edition is None):
        parser.print_usage()
        sys.exit(1)

    setup = Setup(args)
    setup.run()

if __name__ == '__main__':
    main()

