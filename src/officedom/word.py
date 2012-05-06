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

from functools import partial
import itertools
import weakref

import pythoncom
import win32com.client

from utils import ReadOnlyList, WrapperObject

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
        self._langs = _Languages(self._app.Languages)

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

    @property
    def languages(self):
        """Collection of languages

        `self` is this application.

        """
        return self._langs


class _Document(WrapperObject):

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
        WrapperObject.__init__(self, doc)
        self.data = _LightDocument(doc)
        self._parent_docs = weakref.proxy(doc_list)

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
        self._raw_obj.Close(*args, **kwargs)
        self._parent_docs.remove(self)

    def save(self):
        """Save this document.

        `self` is this word document.
        The method synchronizes data to the underlying COM object before
        saving.

        """
        self.data.sync(self._raw_obj)
        self._raw_obj.Save()

    def save_as(self, *args, **kwargs):
        """Save this document to the given file.

        `self` is this word document.
        Positional and keyword arguments are the same as those accepted
        by the corresponding method in word DOM API.
        The method synchronizes data to the underlying COM object before
        saving.

        """
        self.data.sync(self._raw_obj)
        self._raw_obj.SaveAs(*args, **kwargs)


class _Documents(ReadOnlyList):

    """Collection of documents"""

    def __init__(self, docs):
        """Create a collection of documents.

        `self` is this collection of documents.
        `docs` are the COM objects representing documents.

        """
        ReadOnlyList.__init__(self, docs, partial(_Document, self))

    def add(self, *args, **kwargs):
        """Add a new empty document.

        `self` is this collection of documents.
        Positional and keyword arguments are the same as those accepted
        by the corresponding method in word DOM API.
        The method returns the new document.

        """
        self._wrapper_list.append(
            _Document(self, self._raw_list.Add(*args, **kwargs)))
        return self._wrapper_list[-1]

    def close(self, *args, **kwargs):
        """Close all documents.

        `self` is this collection of documents.
        Positional and keyword arguments are the same as those accepted
        by the corresponding method in word DOM API.

        """
        self._raw_list.Close(*args, **kwargs)
        self._wrapper_list = []

    def open(self, file_name, *args, **kwargs):
        """Open the given document file and return it.

        `self` is this collection of documents.
        `file_name` is the file to open.
        Positional and keyword arguments are the same as those accepted
        by the corresponding method in word DOM API.
        If the document is already open, return it. Otherwise open the
        document, add it to this collection, and return it.

        """
        new_doc = self._raw_list.Open(file_name, *args, **kwargs)
        # Check if the document is already open.
        try:
            return self._get_wrapper(new_doc)
        # The requested document isn't open, open and return it.
        except ValueError:

            self._wrapper_list.append(_Document(self, new_doc))
            return self._wrapper_list[-1]

    def remove(self, doc):
        """Remove the given document from this collection.

        `self` is this collection of documents.
        `doc` is the document to remove.
        The method isn't intended for direct manipulation by clients.

        """
        self._wrapper_list.remove(doc)

    def save(self, *args, **kwargs):
        """Save all documents.

        `self` is this collection of documents.
        Positional and keyword arguments are the same as those accepted
        by the corresponding method in word DOM API.

        """
        self._raw_list.Save(*args, **kwargs)


class _Language(WrapperObject):

    """language information"""

    def __init__(self, lang):
        """Create a language.

        `self` is this language.
        `lang` is the underlying COM object representing the language.

        """
        WrapperObject.__init__(self, lang)

    def __getattr__(self, name):
        """Support language properties.

        `self` is this language.
        `name` is the attribute.

        """
        prop_map = {"id": "ID", "name": "Name", "name_local": "NameLocal"}
        prop_names = prop_map.iterkeys()

        if name in prop_names:
            return getattr(self._raw_obj, prop_map[name])

        raise AttributeError()


class _Languages(ReadOnlyList):

    """Collection of languages"""

    def __init__(self, langs):
        """Create a collection of languages.

        `self` is this collection of languages.
        `docs` are the COM objects representing languages.

        """
        ReadOnlyList.__init__(self, langs, _Language)
        
    def Item(self, index):
        """Return the language at the specified index.

        `self` is this collection of languages.
        `index` is the language ID, name, or local name.

        """
        return self._get_wrapper(self._raw_list.Item(index))

    def _get_wrapper(self, raw_obj):
        """Return the wrapper language for the given raw one.

        `self` is this collection of languages.
        `raw_obj` is the raw language to get whose wrapper.
        The method raises a ValueError if no language wraps the given
        raw one.

        """
        # Check if the language is wrapped.
        for cur_lang in self._wrapper_list:
            # Comparison is based on ID's.
            if cur_lang.raw_obj.ID == raw_obj.ID:
                return cur_lang

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
        self.active_writing_style = {}

        for lang in doc.Application.Languages:
            try:  # sorry, no clean way to know available languages
                entry_maker = partial(
                    self._style_entry, lang, doc.ActiveWritingStyle(lang.ID))
            except pythoncom.com_error:
                pass
            else:
                self.active_writing_style.update(
                    itertools.imap(entry_maker, ["ID", "Name", "NameLocal"]))

    def __eq__(self, other):
        """Test if the two documents have the same content.

        `self` is this word document.
        `other` is the other word document.

        """
        # All public data attributes should be equal for the equality
        # test to succeed.
        attr_names = dir(self)

        for cur_attr in attr_names:
            if not cur_attr.startswith('_'):  # Exclude private attributes.

                attr_val = getattr(self, cur_attr)

                if not callable(attr_val) and \
                    attr_val != getattr(other, cur_attr):
                    return False

        return True

    def __ne__(self, other):
        """Test if the two documents have different content.

        `self` is this word document.
        `other` is the other word document.

        """
        return not self == other

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

        # The active writing style property can't be updated.

    @staticmethod
    def _style_entry(lang, style, attr):
        """Return a writing style tuple.

        `lang` is the language to work on.
        `style` is the active writing style of the language.
        `attr` is the key language attribute.
        The method creates and returns a tuple of the specified language
        attribute and its active writing style.

        """
        return getattr(lang, attr), style
