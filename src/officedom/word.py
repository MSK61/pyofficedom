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
import weakref

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

    def quit(self, *args, **kwargs):
        """Close this word application.

        `self` is this application.
        Positional and keyword arguments are the same as those accepted
        by the corresponding method in word DOM API.

        """
        self._app.Quit(*args, **kwargs)

    @property
    def documents(self):
        """Collection of open documents

        `self` is this application.

        """
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
        self._parent_docs = weakref.proxy(doc_list)
        self._doc = doc

    # context manager support
    def __enter__(self):
        """Setup a context for this word document.

        `self` is this word document.

        """
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        """Close this word document.

        `self` is this word document.
        `exc_type` is the type of raised exception if one was raised or
                   None otherwise.
        `exc_value` is the raised exception if one was raised or None
                    otherwise.
        `traceback` is the traceback when the exception occurred, if
                    any, or None otherwise.

        """
        self.close()

    def close(self, *args, **kwargs):
        """Close this document.

        `self` is this word document.
        Positional and keyword arguments are the same as those accepted
        by the corresponding method in word DOM API.
        The method notifies the parent document list about closure.

        """
        self._doc.Close(*args, **kwargs)
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
        Positional and keyword arguments are the same as those accepted
        by the corresponding method in word DOM API.
        The method synchronizes data to the underlying COM object before
        saving.

        """
        self.data.sync(self._doc)
        self._doc.SaveAs(*args, **kwargs)

    @property
    def raw_doc(self):
        """Wrapped COM document object

        `self` is this word document.
        The property isn't intended for direct manipulation by clients.

        """
        return self._doc


class _Documents:

    """Collection of documents"""

    def __init__(self, docs):
        """Create a collection of documents.

        `self` is this collection of documents.
        `docs` are the COM objects representing documents.

        """
        self._raw_docs = docs
        self._docs = map(functools.partial(_Document, self), docs)

    def __getattr__(self, name):
        """Support immutable list operations.

        `self` is this collection of documents.
        `name` is the attribute.

        """
        if name in ["count", "index", "__len__", "__iter__", "__reversed__",
                    "__contains__"]:
            return getattr(self._docs, name)

        raise AttributeError()

    def __getitem__(self, key):
        """Support integer and string indices.

        `self` is this collection of documents.
        `key` is document name/index to look up.

        """
        try:
            return self._docs[key]  # integer indices
        except TypeError:  # string indices
            return self._get_doc_obj(self._raw_docs(key))

    def add(self, *args, **kwargs):
        """Add a new empty document.

        `self` is this collection of documents.
        Positional and keyword arguments are the same as those accepted
        by the corresponding method in word DOM API.
        The method returns the new document.

        """
        self._docs.append(_Document(self, self._raw_docs.Add(*args, **kwargs)))
        return self._docs[-1]

    def close(self, *args, **kwargs):
        """Close all documents.

        `self` is this collection of documents.
        Positional and keyword arguments are the same as those accepted
        by the corresponding method in word DOM API.

        """
        self._raw_docs.Close(*args, **kwargs)
        self._docs = []

    def open(self, file_name, *args, **kwargs):
        """Open the given document file and return it.

        `self` is this collection of documents.
        `file_name` is the file to open.
        Positional and keyword arguments are the same as those accepted
        by the corresponding method in word DOM API.
        If the document is already open, return it. Otherwise open the
        document, add it to this collection, and return it.

        """
        new_doc = self._raw_docs.Open(file_name, *args, **kwargs)
        # Check if the document is already open.
        try:
            return self._get_doc_obj(new_doc)
        # The requested document isn't open, open and return it.
        except ValueError:

            self._docs.append(_Document(self, new_doc))
            return self._docs[-1]

    def remove(self, doc):
        """Remove the given document from this collection.

        `self` is this collection of documents.
        `doc` is the document to remove.
        The method isn't intended for direct manipulation by clients.

        """
        self._docs.remove(doc)

    def save(self, *args, **kwargs):
        """Save all documents.

        `self` is this collection of documents.
        Positional and keyword arguments are the same as those accepted
        by the corresponding method in word DOM API.

        """
        self._raw_docs.Save(*args, **kwargs)

    def _get_doc_obj(self, raw_doc):
        """Return the wrapper document for the given raw document.

        `self` is this collection of documents.
        `raw_doc` is the raw document to get whose wrapper.
        The method raises a ValueError if no document wraps the given
        raw document.

        """
        # Check if the document is wrapped.
        for cur_doc in self._docs:
            if cur_doc.raw_doc == raw_doc:
                return cur_doc

        raise ValueError()


class _LightDocument:

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

    def __eq__(self, other):
        """Test if the two documents have the same content.

        `self` is this word document.
        `other` is the other word document.

        """
        return self.active_theme == other.active_theme

    def __ne__(self, other):
        """Test if the two documents have different content.

        `self` is this word document.
        `other` is the other word document.

        """
        return self.active_theme != other.active_theme

    def sync(self, doc):
        """Update the underlying COM object.

        `self` is this word document.
        `doc` is the underlying COM object to update.
        The method isn't intended for direct manipulation by clients.

        """
        if self.active_theme == NO_OBJ:
            doc.RemoveTheme()
        else:
            doc.ApplyTheme(active_theme)
