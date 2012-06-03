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

from functools import partial
import os
from os.path import abspath, join
import shutil
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
        Verify that the application is quitted  upon context exit.

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
        Open the document again and verify that the theme was changed.

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
            with app.documents.open(out_file) as doc:
                self.assertEqual(doc.data, doc_data)

    def test_tmpl(self):
        """Test changing templates.

        `self` is this test case.
        Load a document, change its template, and then save the document
        again.
        Close the document.
        Open the document again and verify that the template was
        changed.
        Revert the document back to its original template.
        Close the document.
        Open the document again and verify that the template was
        reverted.

        """
        test_doc = "test.doc"
        new_tmpl = "Elegant Letter.dot"
        out_file = join(self._fixture.out_dir, test_doc)
        with Application() as app:

            doc = app.documents.open(join(self._fixture.data_dir, test_doc))
            self.assertEqual(
                doc.attached_template, app.normal_template.full_name)
            # Modify and save the document.
            doc.attached_template = new_tmpl
            doc_data = doc.data
            doc.save_as(out_file)
            doc.close()
            # Reopen and validate the document.
            doc = app.documents.open(out_file)
            self.assertEqual(doc.data, doc_data)
            # Revert the template and save the document.
            doc.attached_template = app.normal_template
            doc_data = doc.data
            doc.close(constants.wdSaveChanges)
            # Reopen and validate the document.
            with app.documents.open(out_file) as doc:
                self.assertEqual(doc.data, doc_data)


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
        """Verify that no cycles exist in the document tree.

        `self` is this test case.
        Load a document.
        Verify that the document object has no cyclic relations
        throughout its object hierarchy.

        """
        test_doc = "test.doc"
        encoding = "utf-8"
        objref_path = ".//*[@objref]"
        with Application() as app:
            with app.documents.open(
                join(self._fixture.data_dir, test_doc)) as doc:
                self.assertIsNone(xml.etree.ElementTree.fromstring(
                    pyxser.serialize(doc.data, encoding)).find(objref_path))

    def test_context(self):
        """Test context manager features of documents.

        `self` is this test case.
        Use an document as the context manager of with statement.
        Verify that the document is closed appropriately upon context
        exit.

        """
        test_doc = "test.doc"
        with Application() as app:

            doc = app.documents.open(join(self._fixture.data_dir, test_doc))
            doc.close = MagicMock(wraps=doc.close)
            with doc:
                pass
            doc.close.assert_called_once_with()

    def test_doc_col(self):
        """Test sequence operations on documents.

        `self` is this test case.
        Verify that typical immutable sequence operations work for documents.

        """
        path_creator = partial(join, self._fixture.data_dir)
        test_doc = "test.doc"
        test_doc2 = "a.doc"
        with Application() as app:

            self.assertEqual(len(app.documents), 0)
            self.assertFalse(app.documents)
            doc = app.documents.open(path_creator(test_doc))
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
            self.assertEqual(
                app.documents.open(path_creator(test_doc2)), app.documents[1])
            self.assertEqual(app.documents[test_doc], doc)  # indexing by name
            self.assertEqual(app.documents.index(doc), 0)
            self.assertEqual(app.documents[::2][0], doc)  # extended slicing
            reversed_docs = list(reversed(app.documents))
            self.assertSequenceEqual(
                app.documents, list(reversed(reversed_docs)))
            self.assertIn(doc, app.documents)

            for cur_doc in app.documents:
                pass

            app.documents.close()

    def test_lang_style(self):
        """Test language writing styles.

        `self` is this test case.
        Load a document and verify its language writing style.

        """
        test_doc = "test.doc"
        lang_style = "grammar only"
        with Application() as app:
            with app.documents.open(
                join(self._fixture.data_dir, test_doc)) as doc:

                self.assertEqual(doc.data.active_writing_style[
                    constants.wdEnglishUS], lang_style)
                lang = app.languages.Item(constants.wdEnglishUS)
                self.assertEqual(
                    doc.data.active_writing_style[lang.name], lang_style)
                self.assertEqual(
                    doc.data.active_writing_style[lang.name_local], lang_style)

    def test_multi_open(self):
        """Test opening the same document several times.

        `self` is this test case.
        Load a document twice.
        Verify that the document is opened only once.

        """
        test_doc = "test.doc"
        in_file = join(self._fixture.data_dir, test_doc)
        with Application() as app:
            with app.documents.open(in_file) as doc:

                self.assertEqual(len(app.documents), 1)
                self.assertEqual(app.documents.open(in_file), doc)
                self.assertEqual(len(app.documents), 1)

    def test_name_prop(self):
        """Test document name property.

        `self` is this test case.
        Verify that the document name property references the document
        file name.

        """
        test_doc = "test.doc"
        with Application() as app:
            with app.documents.open(
                join(self._fixture.data_dir, test_doc)) as doc:
                self.assertEqual(doc.name, test_doc)


class TmplChangeTest(TestCase):

    """Test case for changing template content and properties"""

    def __init__(self, test_func):
        """Create a template change test.

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

    def test_auto_txt(self):
        """Test changing autoText.

        `self` is this test case.
        Load a template, Add autoText entries to it, and then save the
        template again.
        Close the template.
        Open the template again and verify that the autoText entries
        were updated.

        """
        test_tmpl = "test.dot"
        shutil.copy(
            join(self._fixture.data_dir, test_tmpl), self._fixture.out_dir)
        out_tmpl = join(self._fixture.out_dir, test_tmpl)
        new_entries = {"dear": "honey"}
        with Application() as app:

            doc = app.documents.open(
                out_tmpl, Format=constants.wdOpenFormatTemplate)
            tmpl_data = app.templates[doc.attached_template].data
            self.assertFalse(tmpl_data.auto_text_entries)
            # Modify and save the template.
            tmpl_data.auto_text_entries = new_entries
            app.templates[doc.attached_template].save()
            doc.close()
            # Reopen and validate the template.
            with app.documents.open(
                out_tmpl, Format=constants.wdOpenFormatTemplate) as doc:
                self.assertEqual(app.templates[doc.attached_template].data, tmpl_data)


class TmplTest(TestCase):

    """Test case for template properties and operations"""

    def __init__(self, test_func):
        """Create a template test.

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
        """Verify that no cycles exist in the template tree.

        `self` is this test case.
        Verify that the normal template object has no cyclic relations
        throughout its object hierarchy.

        """
        encoding = "utf-8"
        objref_path = ".//*[@objref]"
        with Application() as app:
            self.assertIsNone(
                xml.etree.ElementTree.fromstring(pyxser.serialize(
                    app.normal_template.data, encoding)).find(objref_path))

    def test_doc_rel(self):
        """Verify that relations between documents and templates.

        `self` is this test case.
        Verify that loading templates/documents reflect in each other
        correctly.

        """
        test_doc = "test.dot"
        path_creator = partial(join, self._fixture.data_dir)
        tmpl = "Elegant Letter.dot"
        pure_doc = "test.doc"
        with Application() as app:

            # Open a template as a document and verify that the document
            # collection is updated.
            self.assertFalse(app.documents)
            with app.normal_template.open_as_document():
                self.assertEqual(len(app.documents), 1)
            # Open a document referencing a new template and verify that
            # the template collection is updated.
            # Note that the normal template is always loaded upon
            # starting the word application.
            self.assertEqual(len(app.templates), 1)
            with app.documents.open(path_creator(test_doc)):
                self.assertEqual(len(app.templates), 2)
            # After closing the document, the number of referenced
            # templates falls back.
            self.assertEqual(len(app.templates), 1)
            # Create a new document referencing a new template and
            # verify that the template collection is updated.
            with app.documents.add(tmpl):
                self.assertEqual(len(app.templates), 2)
            # After closing the document, the number of referenced
            # templates falls back.
            self.assertEqual(len(app.templates), 1)
            # Associate a new template to a document and verify that the
            # template collection is updated.
            doc = app.documents.add()
            self.assertEqual(len(app.templates), 1)
            doc.attached_template = tmpl
            self.assertEqual(len(app.templates), 2)
            # After closing all the documents, the number of referenced
            # templates falls back.
            app.documents.close(constants.wdDoNotSaveChanges)
            self.assertEqual(len(app.templates), 1)

    def test_multi_open(self):
        """Test opening the same template in several ways.

        `self` is this test case.
        Load a template twice.
        Verify that the template is opened only once.

        """
        with Application() as app:
            with app.normal_template.open_as_document() as tmpl:

                self.assertEqual(len(app.documents), 1)
                self.assertEqual(app.documents.open(
                    app.normal_template.full_name,
                    Format=constants.wdOpenFormatTemplate), tmpl)
                self.assertEqual(len(app.documents), 1)

    def test_normal(self):
        """Test auto-loading of the normal template.

        `self` is this test case.
        Verify that the normal template is loaded upon application
        startup.

        """
        with Application() as app:
            self.assertIn(app.normal_template, app.templates)

    def test_tmpl_col(self):
        """Test sequence operations on template.

        `self` is this test case.
        Verify that typical immutable sequence operations work for
        templates.

        """
        test_tmpl = "test.dot"
        with Application() as app:

            self.assertEqual(len(app.templates), 1)
            self.assertTrue(app.templates)
            # typical indexing operations
            self.assertEqual(app.templates[-1], app.normal_template)
            # slicing
            self.assertEqual(app.templates[0:1][0], app.normal_template)
            self.assertEqual(app.templates[0:][0], app.normal_template)
            self.assertEqual(app.templates[:1][0], app.normal_template)
            self.assertEqual(app.templates[:][0], app.normal_template)
            self.assertEqual(app.templates.count(app.normal_template), 1)
            # Verify that opening new templates doesn't alter the order
            # of already open templates in the template list.
            with app.documents.open(
                join(self._fixture.data_dir, test_tmpl),
                Format=constants.wdOpenFormatTemplate) as tmpl:

                self.assertEqual(
                    tmpl.attached_template, app.templates[1].full_name)
                # indexing by name
                self.assertEqual(app.templates[app.normal_template.full_name],
                                 app.normal_template)
                self.assertEqual(app.templates.index(app.normal_template), 0)
                # extended slicing
                self.assertEqual(app.templates[::2][0], app.normal_template)
                reversed_tmpls = list(reversed(app.templates))
                self.assertSequenceEqual(
                    app.templates, list(reversed(reversed_tmpls)))

                for cur_tmpl in app.templates:
                    pass


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
