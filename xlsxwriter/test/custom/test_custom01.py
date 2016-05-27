###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from datetime import datetime
from ..helperfunctions import _xml_to_list
from ...custom_properties import CustomProperties


class TestAssembleCustomProperties(unittest.TestCase):
    """
    Test assembling a complete CustomProperties file.

    """
    def test_assemble_xml_file(self):
        """Test writing an Core file."""
        self.maxDiff = None

        fh = StringIO()
        custom = CustomProperties()
        custom._set_filehandle(fh)

        custom_properties = {
            'Test123': 'Jonas',
        }

        custom._set_custom_properties(custom_properties)

        custom._assemble_xml_file()

        exp = _xml_to_list("""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
                  <property name="Test123" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2"><vt:lpwstr>Jonas</vt:lpwstr></property>
                </Properties>
                """)

        got = _xml_to_list(fh.getvalue())

        self.assertEqual(got, exp)
