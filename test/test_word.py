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

import os
from os.path import abspath, join
from shutil import rmtree
import unittest
from unittest import TestCase
import xml.etree.ElementTree

from mock import MagicMock
import pyxser

import Fixture
from officedom.word import Application, constants, NO_OBJ

class AppContextTest(TestCase):

    """Test case for using application objects in a context"""

    def test(self):
        """Test context manager features of application objects.

        `self` is this test case.
        Use an application object as the context manager of with
        statement.
        Verify the application is quitted  upon context exit.

        """
        app = Application()
        app.quit = MagicMock(wraps=app.quit)
        with app:
            pass
        app.quit.assert_called_once_with()


class DocChangeTest(TestCase):

    """Test case for changing document content and properties"""

    def __init__(self, test_func):
        """Create a document change test.

        `self` is this test case.
        `test_func` is the test function to run.

        """
        TestCase.__init__(self, test_func)
        self._fixture = _WorkDirFixture()

    def setUp(self):
        """Prepare for the test.

        `self` is this test case.

        """
        self._fixture.setUp()

    def tearDown(self):
        """Delete the test working directory.

        `self` is this test case.

        """
        self._fixture.tearDown()

    def test_theme(self):
        """Test changing themes.

        `self` is this test case.
        Load a document, change its active theme, and then save the
        document again.
        Close the document.
        Open the document again and verify the theme was changed.

        """
        test_doc = "test.doc"
        out_file = join(self._fixture.out_dir, test_doc)
        with Application() as app:

            doc = app.documents.open(join(self._fixture.data_dir, test_doc))
            doc_data = doc.data
            self.assertNotEqual(doc_data.active_theme, NO_OBJ)
            # Modify and save the document.
            doc_data.active_theme = NO_OBJ
            doc.save_as(out_file)
            doc.close()
            # Reopen and validate the document.
            self.assertEqual(app.documents.open(out_file).data, doc_data)


class DocTest(TestCase):

    """Test case for document properties and operations"""

    def __init__(self, test_func):
        """Create a document test.

        `self` is this test case.
        `test_func` is the test function to run.

        """
        TestCase.__init__(self, test_func)
        self._fixture = _WorkDirFixture()

    def setUp(self):
        """Prepare for the test.

        `self` is this test case.

        """
        self._fixture.setUp()

    def tearDown(self):
        """Delete the test working directory.

        `self` is this test case.

        """
        self._fixture.tearDown()

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

            doc_data = app.documents.open(
                join(self._fixture.data_dir, test_doc)).data
            self.assertIsNone(xml.etree.ElementTree.fromstring(
                pyxser.serialize(doc_data, encoding)).find(objref_path))

    def test_context(self):
        """Test context manager features of documents.

        `self` is this test case.
        Use an document as the context manager of with statement.
        Verify the document is closed appropriately upon context
        exit.

        """
        test_doc = "test.doc"
        with Application() as app:

            doc = app.documents.open(join(self._fixture.data_dir, test_doc))
            doc.close = MagicMock(wraps=doc.close)
            with doc:
                pass
            doc.close.assert_called_once_with()

    def test_doc_seq(self):
        """Test sequence operations on documents.

        `self` is this test case.
        Verify typical immutable sequence operations work for documents.

        """
        test_doc = "test.doc"
        test_doc2 = "a.doc"
        with Application() as app:

            doc = app.documents.open(join(self._fixture.data_dir, test_doc))
            self.assertEqual(len(app.documents), 1)
            # typical indexing operations
            self.assertEqual(app.documents[-1], doc)
            # slicing
            self.assertEqual(app.documents[0:1][0], doc)
            self.assertEqual(app.documents[0:][0], doc)
            self.assertEqual(app.documents[:1][0], doc)
            self.assertEqual(app.documents[:][0], doc)
            self.assertEqual(app.documents.count(doc), 1)
            # Verify that opening new documents doesn't alter the order
            # of already open documents in the document list.
            self.assertEqual(app.documents.open(
                join(self._fixture.data_dir, test_doc2)), app.documents[1])
            self.assertEqual(app.documents[test_doc], doc)  # indexing by name
            self.assertEqual(app.documents.index(doc), 0)
            self.assertEqual(app.documents[::2][0], doc)  # extended slicing
            reversed_docs = list(reversed(app.documents))
            self.assertSequenceEqual(
                app.documents, list(reversed(reversed_docs)))
            self.assertIn(doc, app.documents)

            for cur_doc in app.documents:
                pass

    def test_lang_style(self):
        """Test language writing styles.

        `self` is this test case.
        Load a document and verify its language writing style.

        """
        test_doc = "test.doc"
        orig_style = "Grammar Only"
        with Application() as app:

            doc = app.documents.open(join(self._fixture.data_dir, test_doc))
            self.assertEqual(doc.data.active_writing_style[
                constants.wdEnglishUS], orig_style)

    def test_multi_open(self):
        """Test opening the same document several times.

        `self` is this test case.
        Load a document twice.
        Verify that the document is opened only once.

        """
        test_doc = "test.doc"
        in_file = join(self._fixture.data_dir, test_doc)
        with Application() as app:

            app.documents.open(in_file)
            self.assertEqual(len(app.documents), 1)
            app.documents.open(in_file)
            self.assertEqual(len(app.documents), 1)


class _WorkDirFixture(Fixture.Fixture):

    """Fixture for creating and cleaning up working directories"""

    # input directory
    _DATA_DIR_NAME = "data"

    data_dir = abspath(_DATA_DIR_NAME)

    # result directory
    _OUT_DIR_NAME = "res"

    out_dir = abspath(_OUT_DIR_NAME)

    def setUp(self):
        """Create the working directory for the test.

        `self` is this fixture.
        The old working directory(if any) is removed first.

        """
        if os.path.exists(self.out_dir):  # Clean stale test results(if any).
            rmtree(self.out_dir)

        os.mkdir(self.out_dir)

    def tearDown(self):
        """Delete the test working directory.

        `self` is this fixture.

        """
        rmtree(self.out_dir)

def main():
    """entry point for running test in this module"""
    unittest.main()

if __name__ == '__main__':
    main()
