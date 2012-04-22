# -*- coding: utf-8 -*-

"""tests word DOM"""

############################################################
#
# Copyright 2012 Mohammed El-Afifi
# This file is part of pyofficedom.
#
# pyofficedom is free software: you can redistribute it and/or modify
# it under the terms of the GNU Lesser General Public License as
# published by the Free Software Foundation, either version 3 of the
# License, or (at your option) any later version.
#
# pyofficedom is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU Lesser General Public License for more details.
#
# You should have received a copy of the GNU Lesser General Public
# License along with pyofficedom.  If not, see
# <http://www.gnu.org/licenses/>.
#
# program:      python office DOM
#
# file:         test_load_save.py
#
# function:     word DOM tests
#
# description:  tests word DOM relations
#
# author:       Mohammed Safwat (MS)
#
# environment:  ActiveState Komodo IDE, version 6.1.3, build 66534,
#               windows xp professional
#
# notes:        This is a private program.
#
############################################################

import mock
from officedom.word import Application, NO_OBJ
import os
from os.path import abspath, join
import pyxser
import shutil
from shutil import rmtree
import unittest
from unittest import TestCase
import xml.etree.ElementTree

class AppContextTest(TestCase):

    """Test case for using application objects in a context"""

    def test(self):
        """Test context manager features of application objects.

        `self` is this test case.
        Use an application object as the context manager of with
        statement.
        Verify the application is quitted appropriately upon context
        exit.

        """
        app = Application()
        app.quit = mock.MagicMock(wraps=app.quit)
        with app:
            pass
        app.quit.assert_called_once_with()


class DocTest(TestCase):

    """Test case for document properties and operations"""

    # input directory
    _DATA_DIR_NAME = "data"

    _data_dir = abspath(_DATA_DIR_NAME)

    # result directory
    _OUT_DIR_NAME = "res"

    _out_dir = abspath(_OUT_DIR_NAME)

    def setUp(self):
        """Create the working directory for the test.

        `self` is this test case.
        The old working directory(if any) is removed first.

        """
        # Clean stale test results(if any).
        if os.path.exists(self._out_dir):
            rmtree(self._out_dir)

        os.mkdir(self._out_dir)

    def tearDown(self):
        """Delete the test working directory.

        `self` is this test case.

        """
        rmtree(self._out_dir)

    def test_acyclic(self):
        """Verify no cycles exist in the document tree.

        `self` is this test case.
        Load a document.
        Verify the document object has no cyclic relations throughout
        its object hierarchy.

        """
        test_doc = "test.doc"
        encoding = "utf-8"
        objref_path = ".//*[@objref]"
        with Application() as app:

            doc_data = app.documents.open(join(self._data_dir, test_doc)).data
            self.assertIsNone(xml.etree.ElementTree.fromstring(
                pyxser.serialize(doc_data, encoding)).find(objref_path))

    def test_change_doc(self):
        """Test document consistency across loading and saving.

        `self` is this test case.
        Load a document, change some properties, and then save the
        changes.
        Close the document.
        Open the document again and verify the changed properties were
        saved.

        """
        test_doc = "test.doc"
        out_file = join(self._out_dir, test_doc)
        with Application() as app:

            doc = app.documents.open(join(self._data_dir, test_doc))
            self.assertNotEqual(doc.data.active_theme, NO_OBJ)
            # Modify and save the document.
            doc.data.active_theme = NO_OBJ
            doc.save_as(out_file)
            doc.close()
            # Reopen and validate the document.
            self.assertEqual(
                app.documents.open(out_file).data.active_theme, NO_OBJ)

    def test_doc_seq(self):
        """Test sequence operations on documents.

        `self` is this test case.
        Verify indexing, slicing, and iteration over documents work.

        """
        test_doc = "test.doc"
        in_file = join(self._data_dir, test_doc)
        with Application() as app:

            doc = app.documents.open(in_file)
            self.assertTrue(doc in app.documents)
            self.assertEqual(app.documents[0], doc)
            self.assertEqual(app.documents[-1], doc)
            self.assertEqual(app.documents[0:1][0], doc)
            self.assertEqual(app.documents[0:][0], doc)
            self.assertEqual(app.documents[:1][0], doc)
            self.assertEqual(app.documents[:][0], doc)
            shutil.copy(in_file, self._out_dir)
            app.documents.open(join(self._out_dir, test_doc))
            self.assertEqual(app.documents[::2][0], doc)
            reversed_docs = list(reversed(app.documents))
            self.assertSequenceEqual(
                app.documents, list(reversed(reversed_docs)))

            for cur_doc in app.documents:
                pass

    def test_multi_open(self):
        """Test opening the same document several times.

        `self` is this test case.
        Load a document twice.
        Verify that the document is opened only once.

        """
        test_doc = "test.doc"
        in_file = join(self._data_dir, test_doc)
        with Application() as app:

            app.documents.open(in_file)
            self.assertEqual(len(app.documents), 1)
            app.documents.open(in_file)
            self.assertEqual(len(app.documents), 1)

def main():
    """entry point for running test in this module"""
    unittest.main()

if __name__ == '__main__':
    main()
