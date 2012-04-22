# -*- coding: utf-8 -*-

"""wraps word DOM"""

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
# file:         word.py
#
# function:     word DOM API
#
# description:  contains all classes of word DOM API
#
# author:       Mohammed Safwat (MS)
#
# environment:  ActiveState Komodo IDE, version 6.1.3, build 66534,
#               windows xp professional
#
# notes:        This is a private program.
#
############################################################

import functools
import itertools
import win32com.client

NO_OBJ = "none"

class Application(object):

    """Word application"""

    def __init__(self):
        """Create a word application.

        `self` is this application.
        Hook to an active word application instance or start a new one
        if no current one is running.

        """
        appCls = "Word.Application"
        self._app = win32com.client.DispatchEx(appCls)
        self._docs = _Documents(self._app.Documents)

    # context manager support
    def __enter__(self):
        """Setup a context for this word application.

        `self` is this application.

        """
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        """Exit this word application.

        `self` is this application.
        `exc_type` is the type of raised exception if one was raised or
                   None otherwise.
        `exc_value` is the raised exception if one was raised or None
                    otherwise.
        `traceback` is the traceback when the exception occurred, if
                    any, or None otherwise.

        """
        self.quit()

    def quit(self):
        """Close this word application.

        `self` is this application.

        """
        self._app.Quit()

    @property
    def documents(self):
        """Collection of open documents"""
        return self._docs


class _Document(object):

    """Word document

    This document class is stateful in the sense that it holds a
    permanent reference(throughout its lifetime) to the underlying
    document COM object.
    """

    def __init__(self, doc_list, doc):
        """Create a word document.

        `self` is this word document.
        `doc_list` is the document collection owning this document. This
                   list is notified upon closing this document.
        `doc` is the underlying COM object representing the document.

        """
        self.data = _LightDocument(doc)
        self._parent_docs = doc_list
        self._doc = doc

    def close(self):
        """Close this document.

        `self` is this word document.

        """
        self._doc.Close()
        self._parent_docs.remove(self)

    def save(self):
        """Save this document.

        `self` is this word document.
        The method synchronizes data to the underlying COM object before
        saving.

        """
        self.data.sync(self._doc)
        self._doc.Save()

    def save_as(self, *args, **kwargs):
        """Save this document to the given file.

        `self` is this word document.
        The method synchronizes data to the underlying COM object before
        saving.

        """
        self.data.sync(self._doc)
        self._doc.SaveAs(*args, **kwargs)

    @property
    def full_name(self):
        """Full name of this document"""
        return self._doc.FullName


class _Documents:

    """Collection of documents"""

    def __init__(self, docs):
        """Create a collection of documents.

        `self` is this collection of documents.
        `docs` are the COM objects representing documents.

        """
        self._raw_docs = docs
        self._docs = []
        self._docs.extend(
            itertools.imap(functools.partial(_Document, self._docs), docs))

    def __getattr__(self, name):
        """Support immutable list operations.

        `self` is this collection of documents.
        `name` is the attribute.

        """
        if name in ["__len__", "__getitem__", "__iter__", "__reversed__",
                    "__contains__"]:
            return getattr(self._docs, name)

        raise AttributeError()

    def open(self, file_name, *args, **kwargs):
        """Open the given document file and return it.

        `self` is this collection of documents.
        `file_name` is the file to open.
        Positional and keyword arguments are the same as those accepted
        by the corresponding method in word DOM API.
        If the document is already open, return it. Otherwise open the
        document, add it to this collection, and return it.
        """
        # Check if the document is already open.
        match_docs = itertools.ifilter(
            lambda doc: doc.full_name == file_name, self._docs)

        for cur_doc in match_docs:
            return cur_doc

        # The requested document isn't open, open and return it.
        self._docs.append(_Document(
            self._docs, self._raw_docs.Open(file_name, *args, **kwargs)))
        return self._docs[-1]


class _LightDocument(object):

    """Lightweight word document

    This document class is stateless; it doesn't wrap holds a reference
    to the underlying document COM object.
    """

    def __init__(self, doc):
        """Create a lightweight word document.

        `self` is this word document.
        `doc` is the underlying COM object representing the document.

        """
        self.active_theme = doc.ActiveTheme

    def sync(self, doc):
        """Update the underlying COM object.

        `self` is this word document.
        `doc` is the underlying COM object to update.

        """
        if self.active_theme == NO_OBJ:
            doc.RemoveTheme()
        else:
            doc.ApplyTheme(active_theme)
