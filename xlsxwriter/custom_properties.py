###############################################################################
#
# CustomProperties - A class for writing the Excel XLSX CustomProperties file.
#
# Copyright 2013-2016, John McNamara, jmcnamara@cpan.org
#

# Standard packages.
from datetime import datetime

# Package imports.
from . import xmlwriter


class CustomProperties(xmlwriter.XMLwriter):
    """
    A class for writing the Excel XLSX CustomProperties file.


    """

    ###########################################################################
    #
    # Public API.
    #
    ###########################################################################

    def __init__(self):
        """
        Constructor.

        """

        super(CustomProperties, self).__init__()

        self.custom_properties = {}

	self.pid_index = 2

    ###########################################################################
    #
    # Private API.
    #
    ###########################################################################

    def _assemble_xml_file(self):
        # Assemble and write the XML file.

        # Write the XML declaration.
        self._xml_declaration()
        self._write_cp_custom_properties()

        for name, data in self.custom_properties.items():
            self._write__property(name, data)

        self._xml_end_tag('Properties')

        # Close the file.
        self._xml_close()

    def _set_custom_properties(self, custom_properties):
        # Set the document properties.
        self.custom_properties = custom_properties


    ###########################################################################
    #
    # XML methods.
    #
    ###########################################################################
    def _write_cp_custom_properties(self):
        # Write the <cp:customProperties> element.

        xmlns = ('http://schemas.openxmlformats.org/officeDocument/2006/' +
                    'custom-properties')
        xmlns_vt = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'

        attributes = [
            ('xmlns', xmlns),
            ('xmlns:vt', xmlns_vt)
        ]

        self._xml_start_tag('Properties', attributes)

    def _write__property(self, name, data):
        self._xml_start_tag('property', attributes=[('name', name), ('fmtid', '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}'), ('pid', str(self.pid_index))])
        self._xml_data_element('vt:lpwstr', data)
        self._xml_end_tag('property')
	self.pid_index += 1
