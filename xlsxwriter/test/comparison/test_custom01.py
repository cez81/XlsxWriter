# -*- coding: utf-8 -*-########################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#


from ..excel_comparsion_test import ExcelComparisonTest
from ...workbook import Workbook


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.maxDiff = None

        filename = 'custom01.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_one_custom_property(self):
        """Test one custom property"""

        workbook = Workbook(self.got_filename)
        workbook.set_custom_properties({'Owner': 'Jonas Östanbäck'})
        worksheet = workbook.add_worksheet()
        worksheet = workbook.add_worksheet()
        worksheet = workbook.add_worksheet()

        workbook.close()

        self.assertExcelEqual()
